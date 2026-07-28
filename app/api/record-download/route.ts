const MAX_REQUEST_SIZE = 100_000;

const text = (value: unknown, maxLength: number) =>
  typeof value === "string" ? value.slice(0, maxLength) : "";

const logoId = (value: unknown) => {
  const allowed = [
    "rookies",
    "gp-2024",
    "gp-2025",
    "gp-2026",
    "guild",
    "world-series",
  ];
  return typeof value === "string" && allowed.includes(value) ? value : null;
};

export async function POST(request: Request) {
  const webAppUrl = process.env.SHEETS_WEB_APP_URL;
  const secret = process.env.SHEETS_INGEST_SECRET;

  if (!webAppUrl || !secret) {
    return Response.json(
      { ok: false, error: "Recording is not configured." },
      { status: 503 },
    );
  }

  try {
    const rawBody = await request.text();
    if (!rawBody || rawBody.length > MAX_REQUEST_SIZE) {
      return Response.json(
        { ok: false, error: "Invalid request size." },
        { status: 400 },
      );
    }

    const input = JSON.parse(rawBody) as Record<string, unknown>;
    const projectType =
      input.projectType === "MY FEATURE"
        ? "MY FEATURE"
        : input.projectType === "モチベーショングラフ"
          ? "モチベーショングラフ"
          : null;

    if (!projectType) {
      return Response.json(
        { ok: false, error: "Invalid project type." },
        { status: 400 },
      );
    }

    const faceItems = Array.isArray(input.faceItems)
      ? input.faceItems.slice(0, 4).map((item, index) => {
          const value =
            item && typeof item === "object"
              ? (item as Record<string, unknown>)
              : {};
          return {
            number: index + 1,
            title: text(value.title, 15),
            description: text(value.description, 100),
            hasPhoto: value.hasPhoto === true,
            rookiesLogoId: logoId(value.rookiesLogoId),
          };
        })
      : [];

    const motivationEpisodes = Array.isArray(input.motivationEpisodes)
      ? input.motivationEpisodes.slice(0, 7).map((item, index) => {
          const value =
            item && typeof item === "object"
              ? (item as Record<string, unknown>)
              : {};
          const motivation =
            typeof value.motivation === "number"
              ? Math.max(-100, Math.min(100, value.motivation))
              : 0;
          return {
            number: index + 1,
            period: text(value.period, 10),
            title: text(value.title, 15),
            description: text(value.description, 50),
            motivation,
            rookiesLogoId: logoId(value.rookiesLogoId),
          };
        })
      : [];

    const logoIds = Array.isArray(input.logoIds)
      ? input.logoIds
          .slice(0, 7)
          .map(logoId)
          .filter((value): value is string => value !== null)
      : [];

    const payload = {
      anonymousId: text(input.anonymousId, 100),
      projectType,
      template: text(input.template, 30),
      documentTitle: text(input.documentTitle, 30),
      itemCount:
        typeof input.itemCount === "number"
          ? Math.max(0, Math.min(7, Math.round(input.itemCount)))
          : 0,
      faceItems,
      motivationEpisodes,
      motivationUpSummary: text(input.motivationUpSummary, 80),
      motivationDownSummary: text(input.motivationDownSummary, 80),
      logoIds,
      secret,
    };

    const upstream = await fetch(webAppUrl, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload),
      redirect: "follow",
    });
    const resultText = await upstream.text();
    let result: { ok?: boolean } = {};

    try {
      result = JSON.parse(resultText) as { ok?: boolean };
    } catch {
      // Apps Script should return JSON. Treat any other response as a failure.
    }

    if (!upstream.ok || !result.ok) {
      return Response.json(
        { ok: false, error: "Spreadsheet recording failed." },
        { status: 502 },
      );
    }

    return Response.json(
      { ok: true },
      { headers: { "Cache-Control": "no-store" } },
    );
  } catch (error) {
    console.error("Download recording failed.", error);
    return Response.json(
      { ok: false, error: "Unable to record the download." },
      { status: 500 },
    );
  }
}
