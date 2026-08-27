const GRAPH_SCOPE = "https://graph.microsoft.com/Mail.Read";
const GRAPH_SEND_SCOPE = "https://graph.microsoft.com/Mail.Send";

function requiredEnvironmentVariable(name: string) {
  const value = process.env[name]?.trim();
  if (!value) throw new Error(`Missing required environment variable: ${name}`);
  return value;
}

export function getOutlookConfig() {
  const tenantId = requiredEnvironmentVariable("OUTLOOK_TENANT_ID");
  const sessionSecret = requiredEnvironmentVariable("OUTLOOK_SESSION_SECRET");
  if (sessionSecret.length < 32) {
    throw new Error("OUTLOOK_SESSION_SECRET must be at least 32 characters.");
  }
  return {
    clientId: requiredEnvironmentVariable("OUTLOOK_CLIENT_ID"),
    clientSecret: requiredEnvironmentVariable("OUTLOOK_CLIENT_SECRET"),
    redirectUri: requiredEnvironmentVariable("OUTLOOK_REDIRECT_URI"),
    sessionSecret,
    authorizeUrl: `https://login.microsoftonline.com/${encodeURIComponent(tenantId)}/oauth2/v2.0/authorize`,
    tokenUrl: `https://login.microsoftonline.com/${encodeURIComponent(tenantId)}/oauth2/v2.0/token`,
    scopes: ["offline_access", GRAPH_SCOPE, GRAPH_SEND_SCOPE],
  };
}
