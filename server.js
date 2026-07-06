const express = require("express");
const Airtable = require("airtable");
const { Document, Packer, Paragraph, HeadingLevel, TextRun } = require("docx");

const app = express();
app.use(express.json({ limit: "5mb" }));
app.use(express.static("."));

const base = new Airtable({ apiKey: process.env.AIRTABLE_API_KEY }).base(
  process.env.AIRTABLE_BASE_ID
);

// -----------------------------------------------------------------------
// /api/config - serves the Anthropic API key to the browser for the
// session. Browser-side calls are required to avoid Vercel's 300-second
// function timeout (the full 5-wave pipeline can run well past that).
// -----------------------------------------------------------------------
app.get("/api/config", (req, res) => {
  res.json({ anthropicApiKey: process.env.ANTHROPIC_API_KEY });
});

// -----------------------------------------------------------------------
// Generic table route factory. Every table gets the same five endpoints:
// POST create, GET list, GET one, PATCH update, DELETE.
// Wrapped in try/catch so failures surface in the browser Network tab
// as JSON instead of hanging or returning HTML error pages.
// -----------------------------------------------------------------------
function registerTableRoutes(routeName, tableName) {
  // CREATE
  app.post(`/api/${routeName}`, async (req, res) => {
    try {
      const record = await base(tableName).create([{ fields: req.body }]);
      res.json({ id: record[0].id, fields: record[0].fields });
    } catch (err) {
      res.status(500).json({ error: err.message });
    }
  });

  // LIST ALL
  app.get(`/api/${routeName}`, async (req, res) => {
    try {
      const records = await base(tableName).select().all();
      res.json(records.map((r) => ({ id: r.id, fields: r.fields })));
    } catch (err) {
      res.status(500).json({ error: err.message });
    }
  });

  // GET ONE
  app.get(`/api/${routeName}/:id`, async (req, res) => {
    try {
      const record = await base(tableName).find(req.params.id);
      res.json({ id: record.id, fields: record.fields });
    } catch (err) {
      res.status(500).json({ error: err.message });
    }
  });

  // UPDATE (patch = merge fields, does not clear unspecified fields)
  app.patch(`/api/${routeName}/:id`, async (req, res) => {
    try {
      const record = await base(tableName).update([
        { id: req.params.id, fields: req.body },
      ]);
      res.json({ id: record[0].id, fields: record[0].fields });
    } catch (err) {
      res.status(500).json({ error: err.message });
    }
  });

  // DELETE
  app.delete(`/api/${routeName}/:id`, async (req, res) => {
    try {
      await base(tableName).destroy([req.params.id]);
      res.json({ deleted: true });
    } catch (err) {
      res.status(500).json({ error: err.message });
    }
  });
}

// Table name here must match the Airtable table name EXACTLY -
// case-sensitive, no trailing spaces (see Gotcha 4).
registerTableRoutes("ideas", "Ideas");
registerTableRoutes("agent-outputs", "AgentOutputs");
registerTableRoutes("blueprints", "Blueprints");
registerTableRoutes("activity-feed", "ActivityFeed");

// -----------------------------------------------------------------------
// Word Doc export - server-side, since this completes in milliseconds
// and doesn't touch Claude at all.
// -----------------------------------------------------------------------
app.get("/api/blueprints/:id/export", async (req, res) => {
  try {
    const record = await base("Blueprints").find(req.params.id);
    const f = record.fields;

    // Blueprints links to Ideas rather than storing the name directly -
    // follow the link to get the idea's name for the title/filename.
    let ideaName = "Execution Blueprint";
    if (f["Idea"] && f["Idea"][0]) {
      const ideaRecord = await base("Ideas").find(f["Idea"][0]);
      ideaName = ideaRecord.fields["Idea Name"] || ideaName;
    }

    const section = (title, body) => [
      new Paragraph({ text: title, heading: HeadingLevel.HEADING_1 }),
      new Paragraph({ children: [new TextRun(body || "")] }),
      new Paragraph({ text: "" }),
    ];

    const doc = new Document({
      sections: [
        {
          properties: {
            page: { size: { width: 12240, height: 15840 } }, // US Letter
          },
          children: [
            new Paragraph({
              text: ideaName,
              heading: HeadingLevel.TITLE,
            }),
            ...section("Executive Briefing", f["Executive Briefing"]),
            ...section("Operational Design", f["Operational Design"]),
            ...section("Financial Model", f["Financial Model"]),
            ...section("Brand Package", f["Brand Package"]),
            ...section("90-Day Launch Playbook", f["Launch Playbook"]),
            ...section("Problem Areas", f["Problem Areas"]),
            ...section("Next 10 Actions", f["Next 10 Actions"]),
          ],
        },
      ],
    });

    const buffer = await Packer.toBuffer(doc);
    const safeName = ideaName.replace(/[^a-z0-9]/gi, "_");

    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    );
    res.setHeader(
      "Content-Disposition",
      `attachment; filename="${safeName}_Blueprint.docx"`
    );
    res.send(buffer);
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`idea-blueprint server running on ${PORT}`));
