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

export const api = onRequest(app);
