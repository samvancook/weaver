import express from "express";
import cors from "cors";
import { onRequest } from "firebase-functions/v2/https";

const app = express();

app.use(cors({ origin: true }));
app.use(express.json());

app.get("/health", (_req, res) => {
  res.json({
    ok: true,
    service: "excerpt-review-api"
  });
});

app.get("/bootstrap", (_req, res) => {
  res.json({
    appName: "Excerpt Review Tool",
    sections: [
      { id: "intake", label: "Intake Queue" },
      { id: "duplicates", label: "Duplicate Review" },
      { id: "approved", label: "Approved" },
      { id: "exports", label: "Export Queue" }
    ]
  });
});

app.post("/catalog/validate", (_req, res) => {
  res.status(501).json({
    ok: false,
    error: "Catalog validation is not available on the hosted site yet. Use persisted sheet-backed validation or the local dev environment for live catalog checks."
  });
});

export const api = onRequest(app);
