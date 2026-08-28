import { execFile, spawn } from "node:child_process";
import path from "node:path";
import { promisify } from "node:util";
import { validateRelativeRepositoryPath, WORKSPACE_ROOT } from "@/lib/developer/workspace";

const execFileAsync = promisify(execFile);
const MAX_GIT_OUTPUT_BYTES = 256 * 1024;
export const MAX_DIFF_CONTENT_CHARS = 80_000;
const MAX_HISTORY_COUNT = 20;

async function runGit(args: readonly string[]) {
  const { stdout } = await execFileAsync("git", [...args], {
    cwd: WORKSPACE_ROOT,
    encoding: "utf8",
    maxBuffer: MAX_GIT_OUTPUT_BYTES,
    windowsHide: true,
    env: { ...process.env, GIT_PAGER: "cat", GIT_TERMINAL_PROMPT: "0" },
  });
  return stdout;
}

async function runGitLimited(args: readonly string[], maxBytes: number) {
  return new Promise<{ output: string; truncated: boolean }>((resolve, reject) => {
    const child = spawn("git", [...args], {
      cwd: WORKSPACE_ROOT,
      windowsHide: true,
      shell: false,
      env: { ...process.env, GIT_PAGER: "cat", GIT_TERMINAL_PROMPT: "0" },
      stdio: ["ignore", "pipe", "pipe"],
    });
    const chunks: Buffer[] = [];
    let bytes = 0;
    let truncated = false;
    let stderr = "";
    child.stdout.on("data", (chunk: Buffer) => {
      if (truncated) return;
      const remaining = maxBytes - bytes;
      if (chunk.length >= remaining) {
        chunks.push(chunk.subarray(0, Math.max(remaining, 0)));
        bytes = maxBytes;
        truncated = true;
        child.kill();
      } else {
        chunks.push(chunk);
        bytes += chunk.length;
      }
    });
    child.stderr.on("data", (chunk: Buffer) => {
      stderr = (stderr + chunk.toString("utf8")).slice(0, 1_000);
    });
    child.on("error", reject);
    child.on("close", (code) => {
      if (code !== 0 && !truncated) {
        reject(new Error(stderr || "Git command failed."));
        return;
      }
      resolve({ output: Buffer.concat(chunks).toString("utf8"), truncated });
    });
  });
}

function statusKind(index: string) {
  if (index === "?" || index === "!") return null;
  return index === " " ? null : index;
}

export async function getRepositoryStatus() {
  const [root, branch, statusOutput] = await Promise.all([
    runGit(["rev-parse", "--show-toplevel"]),
    runGit(["branch", "--show-current"]),
    runGit(["status", "--porcelain=v1", "--no-renames", "-z"]),
  ]);
  const records = statusOutput.split("\0").filter(Boolean);
  const changedRecordCount = records.filter((record) => record[0] !== "!").length;
  const files = records.flatMap((record) => {
    if (record.length < 4) return [];
    const index = record[0];
    const worktree = record[1];
    const filePath = record.slice(3);
    if (index === "!" || !filePath || !validateRelativeRepositoryPath(filePath)) return [];
    return [{
      path: filePath,
      staged: statusKind(index),
      unstaged: statusKind(worktree),
      untracked: index === "?" && worktree === "?",
    }];
  });
  return {
    repositoryRoot: root.trim(),
    repositoryName: path.basename(root.trim()),
    branch: branch.trim() || "(detached HEAD)",
    clean: changedRecordCount === 0,
    files,
    excludedSensitiveFileCount: Math.max(0, changedRecordCount - files.length),
  };
}

export async function getRecentGitHistory(limit: number) {
  const safeLimit = Math.min(Math.max(Math.trunc(limit), 1), MAX_HISTORY_COUNT);
  const output = await runGit([
    "log", `-${safeLimit}`, "--date=iso-strict",
    "--pretty=format:%H%x1f%h%x1f%ad%x1f%s%x1e",
  ]);
  return output.split("\x1e").flatMap((record) => {
    const [hash, shortHash, date, subject] = record.trim().split("\x1f");
    return hash && shortHash && date && subject ? [{ hash, shortHash, date, subject }] : [];
  });
}

export type GitDiffScope = "all" | "staged" | "unstaged";

export async function inspectGitDiff(
  scope: GitDiffScope,
  relativePath: string,
  includeContent: boolean,
) {
  const validatedPath = validateRelativeRepositoryPath(relativePath, true);
  if (validatedPath === null) throw new Error("Invalid or blocked repository path.");
  const repositoryStatus = await getRepositoryStatus();
  const relevantStatus = validatedPath
    ? repositoryStatus.files.filter((file) =>
        file.path === validatedPath || file.path.startsWith(`${validatedPath}/`))
    : repositoryStatus.files;
  const baseArgs = scope === "all" ? ["diff", "HEAD"] :
    scope === "staged" ? ["diff", "--cached"] : ["diff"];
  const changedPathOutput = await runGit([
    ...baseArgs,
    "--name-only",
    "-z",
    ...(validatedPath ? ["--", validatedPath] : []),
  ]);
  const allChangedPaths = changedPathOutput.split("\0")
    .map((value) => value.trim())
    .filter(Boolean);
  const allowedChangedPaths = allChangedPaths
    .filter((value) => Boolean(validateRelativeRepositoryPath(value)));
  const changedPaths = allowedChangedPaths.slice(0, 100);
  if (!changedPaths.length) {
    return {
      scope,
      path: validatedPath || null,
      stat: "",
      changedFiles: "",
      content: includeContent ? "" : null,
      truncated: false,
      excludedSensitiveFileCount: allChangedPaths.length - allowedChangedPaths.length,
      workingTree: {
        files: relevantStatus,
        excludedSensitiveFileCount: repositoryStatus.excludedSensitiveFileCount,
      },
      note: scope === "all" ? "Untracked files are reported by repository status but have no Git diff until tracked." : null,
    };
  }
  const pathArgs = ["--", ...changedPaths];
  const [stat, names, contentResult] = await Promise.all([
    runGit([...baseArgs, "--stat", ...pathArgs]),
    runGit([...baseArgs, "--name-status", ...pathArgs]),
    includeContent
      ? runGitLimited([...baseArgs, "--no-ext-diff", "--unified=3", ...pathArgs], MAX_DIFF_CONTENT_CHARS)
      : Promise.resolve({ output: "", truncated: false }),
  ]);
  return {
    scope,
    path: validatedPath || null,
    stat: stat.trim(),
    changedFiles: names.trim(),
    content: includeContent ? contentResult.output : null,
    truncated: contentResult.truncated || allowedChangedPaths.length > changedPaths.length,
    excludedSensitiveFileCount: allChangedPaths.length - allowedChangedPaths.length,
    workingTree: {
      files: relevantStatus,
      excludedSensitiveFileCount: repositoryStatus.excludedSensitiveFileCount,
    },
    note: scope === "all" ? "Untracked files are reported by repository status but have no Git diff until tracked." : null,
  };
}
