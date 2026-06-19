"use strict";

/* Draft in Voice - Word task-pane logic (deterministic, Stage 1).
   Reads a sections file (your skeleton + your draft-in-voice prose + supervisor notes)
   and inserts a chosen section into Word at the cursor with native styles.
   No AI runs here. It only places what your Claude Code skills already produced. */

var INWORD = false;

/* An embedded sample so the pane is never blank on first open, and so it works
   even before you have loaded a file. Load your own with the "Load sections" button. */
var SAMPLE = {
  title: "Sample: how can AI be governed safely in children's social care?",
  sections: [
    {
      heading: "The value of AI in social care is not what it does or how quickly it does it, but the space it creates for the compassion, principles and time a social worker needs to act.",
      paragraphType: "establishing",
      lead: "AI's worth in this setting is measured by what it frees up for practice, not by raw efficiency.",
      whyItMatters: "This sets the yardstick the rest of the argument uses to judge any AI tool.",
      bullets: [
        { text: "The value of AI is not in what it does or how quickly it does it. It is the space created for the compassion, principles and time that a social worker is provided with to act and intervene.", cite: "(Desouza2020, p.211)" },
        { text: "Why would that financial decision have an ethical consequence? It introduces the idea of the public good within an economic context.", cite: "(Desouza2020, p.206)" }
      ],
      placeholder: "[Write this section in your own words.]",
      draft: "The value of AI in social care, I would argue, is not what it does or how quickly it does it, but the space it creates for the compassion, principles and time a social worker needs to act (Desouza2020, p.211). Seen in this way, its worth is measured by what it frees up for practice rather than by raw efficiency.",
      comments: [
        { text: "This rests on two of your own comments about value. Do you have a source on what an efficiency-first AI actually displaces in practice?" }
      ]
    },
    {
      heading: "Past ICT modernisation in social care destroyed value because it was poorly planned and evaluated internally, so AI needs external evaluation and auditing built in as due diligence.",
      paragraphType: "tension",
      lead: "The lesson of the earlier ICT reform is that internal evaluation let value erode unchecked.",
      whyItMatters: "It shows why an AI governance model cannot rely on self-assessment.",
      bullets: [
        { text: "Value destruction. The previous Labour government's effort to modernise ICT was costly, poorly planned and executed, and a big issue was that it was being evaluated internally.", cite: "(Desouza2020, p.208)" },
        { text: "A system to audit and to preemptively test for biases must be in place as part of the LA's due diligence.", cite: "(Desouza2020, p.211)" }
      ],
      placeholder: "[Write this section in your own words.]",
      comments: [
        { text: "This sentence makes two claims at once, that past ICT destroyed value and that AI now needs external evaluation. Is that one point or two?" }
      ]
    },
    {
      heading: "A workable AI governance model needs to assign roles and concessions that share risk across the network and remove the blame placed on Local Authorities.",
      paragraphType: "synthesis",
      lead: "Sharing risk across the actors is what stops the blame culture reproducing itself.",
      whyItMatters: "This is the move the whole section builds toward, the shape of the solution.",
      bullets: [
        { text: "What are the concessions that remove blame and remove failure for Local Authorities? A conceptual framework that can assign roles and map the interests so that risk is shared across the network.", cite: "(Desouza2020, p.208)" }
      ],
      placeholder: "[Write this section in your own words.]"
    }
  ]
};

if (typeof Office !== "undefined" && Office.onReady) {
  Office.onReady(function (info) {
    INWORD = !!(info && info.host === Office.HostType.Word);
    setStatus(INWORD
      ? "Connected to Word. Put your cursor where you want the text, then Insert."
      : "Preview mode (not inside Word). Open this pane in Word to insert.");
  });
}

document.addEventListener("DOMContentLoaded", function () {
  document.getElementById("file").addEventListener("change", onFile);
  document.getElementById("sampleBtn").addEventListener("click", function () { render(SAMPLE); setStatus("Showing the sample."); });
  render(SAMPLE); // never start blank
});

function setStatus(msg) { var s = document.getElementById("status"); if (s) s.textContent = msg; }

function onFile(e) {
  var f = e.target.files[0]; if (!f) return;
  var r = new FileReader();
  r.onload = function () {
    try { render(JSON.parse(String(r.result))); setStatus("Loaded " + f.name); }
    catch (err) { setStatus("Could not read that file as JSON."); }
  };
  r.readAsText(f); e.target.value = "";
}

function render(data) {
  var sections = (data && data.sections) || [];
  var list = document.getElementById("list");
  list.innerHTML = "";
  if (data && data.title) { var t = document.createElement("div"); t.className = "doctitle"; t.textContent = data.title; list.appendChild(t); }
  if (!sections.length) { list.innerHTML = '<div class="empty">No sections in this file.</div>'; return; }
  sections.forEach(function (sec, i) {
    var card = document.createElement("div"); card.className = "card";
    var hd = document.createElement("div"); hd.className = "cardhead";
    if (sec.paragraphType) { var tag = document.createElement("span"); tag.className = "tag"; tag.textContent = sec.paragraphType; hd.appendChild(tag); }
    var ht = document.createElement("span"); ht.textContent = sec.heading || ("Point " + (i + 1)); hd.appendChild(ht);
    card.appendChild(hd);
    if (sec.draft) { var pv = document.createElement("div"); pv.className = "preview"; pv.textContent = sec.draft.slice(0, 150) + (sec.draft.length > 150 ? "..." : ""); card.appendChild(pv); }
    var btns = document.createElement("div"); btns.className = "btns";
    btns.appendChild(mkBtn("Insert skeleton", function () { insertSkeleton(sec); }));
    if (sec.draft) btns.appendChild(mkBtn("Insert draft", function () { insertDraft(sec); }, "primary"));
    card.appendChild(btns);
    (sec.comments || []).forEach(function (c) {
      var row = document.createElement("div"); row.className = "crow";
      var txt = document.createElement("div"); txt.className = "ctext"; txt.textContent = "Supervisor: " + (c.text || ""); row.appendChild(txt);
      row.appendChild(mkBtn("Comment on selection", function () { commentOnSelection(c.text); }, "sm"));
      card.appendChild(row);
    });
    list.appendChild(card);
  });
}

function mkBtn(label, fn, cls) { var b = document.createElement("button"); b.className = "btn" + (cls ? " " + cls : ""); b.textContent = label; b.onclick = fn; return b; }

function needWord() { if (!INWORD) { setStatus("Open this pane inside Word to insert. You are in preview mode."); return false; } return true; }

function grey(p) { p.font.italic = true; p.font.color = "#888888"; }

function insertSkeleton(sec) {
  if (!needWord()) return;
  Word.run(function (ctx) {
    var last = ctx.document.getSelection().insertParagraph(sec.heading || "Untitled point", Word.InsertLocation.after);
    last.styleBuiltIn = Word.BuiltInStyleName.heading1;
    if (sec.lead) { var l = last.insertParagraph(sec.lead, Word.InsertLocation.after); l.styleBuiltIn = Word.BuiltInStyleName.normal; last = l; }
    if (sec.whyItMatters) { var w = last.insertParagraph("Why it matters: " + sec.whyItMatters, Word.InsertLocation.after); w.styleBuiltIn = Word.BuiltInStyleName.normal; grey(w); last = w; }
    (sec.bullets || []).forEach(function (b) { var bp = last.insertParagraph((b.text || "") + (b.cite ? " " + b.cite : ""), Word.InsertLocation.after); bp.styleBuiltIn = Word.BuiltInStyleName.listBullet; last = bp; });
    var ph = last.insertParagraph(sec.placeholder || "[Write this section in your own words.]", Word.InsertLocation.after); ph.styleBuiltIn = Word.BuiltInStyleName.normal; grey(ph);
    return ctx.sync().then(function () { setStatus("Inserted the skeleton. Write into the grey line."); });
  }).catch(function (e) { setStatus("Insert failed: " + e.message); });
}

function insertDraft(sec) {
  if (!needWord()) return;
  Word.run(function (ctx) {
    var last = ctx.document.getSelection().insertParagraph(sec.heading || "Untitled point", Word.InsertLocation.after);
    last.styleBuiltIn = Word.BuiltInStyleName.heading1;
    (sec.draft || "").split(/\n\n+/).forEach(function (par) {
      if (!par.trim()) return;
      var dp = last.insertParagraph(par.trim(), Word.InsertLocation.after); dp.styleBuiltIn = Word.BuiltInStyleName.normal; last = dp;
    });
    return ctx.sync().then(function () { setStatus("Inserted the draft. This is a starting block to rewrite in your words."); });
  }).catch(function (e) { setStatus("Insert failed: " + e.message); });
}

function commentOnSelection(text) {
  if (!needWord()) return;
  Word.run(function (ctx) {
    var sel = ctx.document.getSelection();
    var canComment = Office.context.requirements && Office.context.requirements.isSetSupported("WordApi", "1.4");
    if (canComment) {
      sel.insertComment(text);
      return ctx.sync().then(function () { setStatus("Comment added to your selection."); });
    }
    var p = sel.insertParagraph("[Supervisor: " + text + "]", Word.InsertLocation.after); grey(p);
    return ctx.sync().then(function () { setStatus("This Word does not support comments here, added it as a note instead."); });
  }).catch(function (e) { setStatus("Select some text in Word first. (" + e.message + ")"); });
}
