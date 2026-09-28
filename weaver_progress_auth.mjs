import { OAuth2Client } from "google-auth-library";

const SERVICE_AUDIENCE = "https://weaver-912447899335.us-central1.run.app";
const EXPORT_SERVICE_ACCOUNT = "weaver-deployer@button-weaver-internal.iam.gserviceaccount.com";
const idTokenVerifier = new OAuth2Client();

function authError(message, statusCode) {
  const error = new Error(message);
  error.statusCode = statusCode;
  return error;
}

export async function verifyProgressExportCaller(req, verifyAdministrator, verifier = idTokenVerifier) {
  const match = String(req.headers.authorization || "").trim().match(/^Bearer\s+(.+)$/i);
  const token = match?.[1]?.trim() || "";
  if (token.split(".").length !== 3) return verifyAdministrator(req);

  let payload;
  try {
    const ticket = await verifier.verifyIdToken({ idToken: token, audience: SERVICE_AUDIENCE });
    payload = ticket.getPayload();
  } catch {
    throw authError("progress_export_token_invalid", 401);
  }
  if (payload?.email !== EXPORT_SERVICE_ACCOUNT || payload.email_verified !== true) {
    throw authError("progress_export_access_denied", 403);
  }
  return { email: payload.email, subject: payload.sub, audience: SERVICE_AUDIENCE };
}
