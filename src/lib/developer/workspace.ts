import { readdir, readFile, realpath, stat } from "node:fs/promises";
import path from "node:path";

const DEFAULT_WORKSPACE_ROOT = "C:\\dev\\jarvis";
export const MAX_FILE_BYTES = 64 * 1024;
export const MAX_TREE_ENTRIES = 300;
export const MAX_TREE_DEPTH = 4;

const DENIED_DIRECTORIES = new Set([
  ".git", ".next", "node_modules", ".ssh", ".aws", ".azure",
  "secrets", ".secrets", "credentials", "certs", "private",
]);
const DENIED_FILE_NAMES = new Set([
  ".npmrc", ".pypirc", ".netrc", "credentials.json", "secrets.json",
  "id_rsa", "id_ed25519",
]);
const DENIED_FILE_EXTENSIONS = new Set([
  ".pem", ".key", ".p12", ".pfx", ".jks", ".keystore", ".secret", ".token",
]);

export const WORKSPACE_ROOT = path.resolve(
  /* turbopackIgnore: true */
  process.env.JARVIS_WORKSPACE_ROOT?.trim() || DEFAULT_WORKSPACE_ROOT,
);

function isWithinRoot(root: string, candidate: string) {
  const relative = path.relative(root, candidate);
  return relative === "" || (!relative.startsWith("..") && !path.isAbsolute(relative));
}

function pathSegments(candidate: string) {
  return candidate.split(/[\\/]+/).filter(Boolean);
}

export function isSensitiveRepositoryPath(relativePath: string) {
  const segments = pathSegments(relativePath);
  if (segments.some((segment) => DENIED_DIRECTORIES.has(segment.toLowerCase()))) {
    return true;
  }
  const fileName = (segments.at(-1) ?? "").toLowerCase();
  return fileName === ".env" || fileName.startsWith(".env.") ||
    DENIED_FILE_NAMES.has(fileName) || DENIED_FILE_EXTENSIONS.has(path.extname(fileName));
}

export function validateRelativeRepositoryPath(value: string, allowEmpty = false) {
  const trimmed = value.trim().replace(/\\/g, "/");
  if (!trimmed) return allowEmpty ? "" : null;
  if (trimmed.length > 500 || path.isAbsolute(trimmed) || trimmed.includes("\0")) return null;
  const resolved = path.resolve(/* turbopackIgnore: true */ WORKSPACE_ROOT, trimmed);
  if (!isWithinRoot(WORKSPACE_ROOT, resolved)) return null;
  const relative = path.relative(WORKSPACE_ROOT, resolved).replace(/\\/g, "/");
  return relative && !isSensitiveRepositoryPath(relative) ? relative : null;
}

async function safeExistingPath(relativePath: string) {
  const root = await realpath(/* turbopackIgnore: true */ WORKSPACE_ROOT);
  const candidate = await realpath(/* turbopackIgnore: true */ path.join(root, relativePath));
  if (!isWithinRoot(root, candidate)) throw new Error("Repository path escapes the workspace.");
  return candidate;
}

export async function readRepositoryTextFile(relativePath: string) {
  const validated = validateRelativeRepositoryPath(relativePath);
  if (!validated) throw new Error("Invalid or blocked repository path.");
  const absolutePath = await safeExistingPath(validated);
  const metadata = await stat(/* turbopackIgnore: true */ absolutePath);
  if (!metadata.isFile()) throw new Error("Repository path is not a file.");
  if (metadata.size > MAX_FILE_BYTES) throw new Error("Repository file exceeds the read limit.");
  const content = await readFile(/* turbopackIgnore: true */ absolutePath);
  if (content.includes(0)) throw new Error("Binary repository files cannot be read.");
  return {
    path: validated,
    size: metadata.size,
    content: content.toString("utf8"),
  };
}

export async function listRepositoryTree(relativePath: string, maxDepth: number) {
  const validated = validateRelativeRepositoryPath(relativePath, true);
  if (validated === null) throw new Error("Invalid or blocked repository path.");
  const absoluteStart = validated ? await safeExistingPath(validated) : await realpath(/* turbopackIgnore: true */ WORKSPACE_ROOT);
  const metadata = await stat(/* turbopackIgnore: true */ absoluteStart);
  if (!metadata.isDirectory()) throw new Error("Repository path is not a directory.");
  const entries: Array<{ path: string; type: "file" | "directory" }> = [];
  let truncated = false;

  async function visit(absoluteDirectory: string, depth: number) {
    if (depth > maxDepth || truncated) return;
    const children = await readdir(/* turbopackIgnore: true */ absoluteDirectory, { withFileTypes: true });
    children.sort((a, b) => a.name.localeCompare(b.name));
    for (const child of children) {
      const absoluteChild = path.join(/* turbopackIgnore: true */ absoluteDirectory, child.name);
      const relativeChild = path.relative(WORKSPACE_ROOT, absoluteChild).replace(/\\/g, "/");
      if (isSensitiveRepositoryPath(relativeChild) || child.isSymbolicLink()) continue;
      if (entries.length >= MAX_TREE_ENTRIES) {
        truncated = true;
        return;
      }
      if (child.isDirectory()) {
        entries.push({ path: relativeChild, type: "directory" });
        await visit(absoluteChild, depth + 1);
      } else if (child.isFile()) {
        entries.push({ path: relativeChild, type: "file" });
      }
    }
  }

  await visit(absoluteStart, 0);
  return { root: validated || ".", entries, truncated };
}

export const BLOCKED_REPOSITORY_PATHS = {
  directories: [...DENIED_DIRECTORIES],
  files: [...DENIED_FILE_NAMES, ".env", ".env.*"],
  extensions: [...DENIED_FILE_EXTENSIONS],
};
