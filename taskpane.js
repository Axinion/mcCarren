"use strict";
// Word document redaction add-in
// Finds and redacts emails, phone numbers, and SSNs
const REDACTION_LABELS = {
    email: "[REDACTED EMAIL]",
    phone: "[REDACTED PHONE]",
    ssn: "[REDACTED SSN]",
};
const HEADER_TEXT = "CONFIDENTIAL DOCUMENT";
// Helper functions for DOM
function getStatusEl() {
    return document.getElementById("status");
}
function getButton() {
    return document.getElementById("redact-btn");
}
Office.onReady(() => {
    const button = getButton();
    if (button) {
        button.addEventListener("click", function () {
            runRedaction().catch(reportError);
        });
    }
});
function updateStatus(msg) {
    const el = getStatusEl();
    if (el) {
        el.textContent = msg;
    }
}
function reportError(error) {
    console.error("Error:", error);
    let msg = "Something went wrong. Please try again.";
    if (typeof error === "string") {
        msg = error;
    }
    else if (error instanceof Error) {
        msg = error.message;
        // Check for Office API debug info
        const err = error;
        if (err.debugInfo) {
            msg += ` (Debug: ${JSON.stringify(err.debugInfo)})`;
        }
        if (err.code) {
            msg += ` [Code: ${err.code}]`;
        }
    }
    else if (error && typeof error === "object") {
        const officeErr = error;
        if (officeErr.debugInfo) {
            msg = `Error: ${officeErr.message || "Unknown"}. Debug: ${JSON.stringify(officeErr.debugInfo)}`;
        }
        else if (officeErr.message) {
            msg = officeErr.message;
        }
    }
    updateStatus(msg);
}
// Find all sensitive patterns in the text
function findSensitivePatterns(text) {
    const results = [];
    // Helper to add unique matches
    function addMatches(regex, kind) {
        const seen = new Set();
        for (const match of text.matchAll(regex)) {
            const val = match[0];
            if (!seen.has(val)) {
                seen.add(val);
                results.push({ value: val, kind });
            }
        }
    }
    // Email regex - catches most common formats
    addMatches(/[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}/g, "email");
    // Phone numbers - handles (555) 123-4567, 555-123-4567, etc.
    addMatches(/\b(?:\+?\d{1,2}[\s.-]?)?(?:\(?\d{3}\)?[\s.-]?)\d{3}[\s.-]?\d{4}\b/g, "phone");
    // SSN format
    addMatches(/\b\d{3}-\d{2}-\d{4}\b/g, "ssn");
    return results;
}
function getStats(matches) {
    const counts = { email: 0, phone: 0, ssn: 0 };
    for (const m of matches) {
        counts[m.kind]++;
    }
    return {
        total: matches.length,
        byKind: counts,
    };
}
function buildMessage(stats, trackingMsg) {
    const parts = [];
    if (stats.byKind.email > 0) {
        parts.push(`${stats.byKind.email} email${stats.byKind.email !== 1 ? "s" : ""}`);
    }
    if (stats.byKind.phone > 0) {
        parts.push(`${stats.byKind.phone} phone${stats.byKind.phone !== 1 ? "s" : ""}`);
    }
    if (stats.byKind.ssn > 0) {
        parts.push(`${stats.byKind.ssn} SSN${stats.byKind.ssn !== 1 ? "s" : ""}`);
    }
    let msg = `Redaction complete. Replaced ${stats.total} item${stats.total !== 1 ? "s" : ""}`;
    if (parts.length > 0) {
        msg += ` (${parts.join(", ")})`;
    }
    if (trackingMsg) {
        msg += `. ${trackingMsg}`;
    }
    return msg;
}
async function setupTrackChanges(context) {
    let note = "";
    // Check if Word API 1.5 is available
    if (!Office.context.requirements.isSetSupported("WordApi", "1.5")) {
        note = "Track Changes not available in this host (skipping tracking).";
        return note;
    }
    try {
        // Try to enable tracking
        context.document.trackRevisions = true;
        await context.sync();
    }
    catch (e) {
        // Might already be enabled or not supported - either way we continue
        const errMsg = e?.message || e?.toString() || "";
        if (errMsg.includes("ApiNotFound") || errMsg.includes("not supported")) {
            note = "Track Changes not available in this host (skipping tracking).";
        }
        // If it's already on, that's fine - we want tracking anyway
    }
    return note;
}
async function insertHeader(context) {
    try {
        const sections = context.document.sections;
        sections.load("items");
        await context.sync();
        if (sections.items.length === 0) {
            return; // No sections, skip
        }
        const firstSection = sections.items[0];
        const header = firstSection.getHeader(Word.HeaderFooterType.primary);
        const headerRange = header.getRange();
        headerRange.load("text");
        await context.sync();
        // Only add if not already there
        if (!headerRange.text.includes(HEADER_TEXT)) {
            header.insertParagraph(HEADER_TEXT, Word.InsertLocation.start);
            await context.sync();
        }
    }
    catch (e) {
        // Header failure shouldn't stop redaction
        console.warn("Header insert failed:", e);
    }
}
async function performRedaction(context, matches) {
    let count = 0;
    for (const match of matches) {
        try {
            // Search for this exact value
            const found = context.document.body.search(match.value, {
                matchCase: false,
                matchWholeWord: false,
            });
            context.load(found, "items");
            await context.sync();
            if (found.items.length === 0) {
                continue; // Not found, skip
            }
            // Replace all occurrences
            found.items.forEach((range) => {
                range.insertText(REDACTION_LABELS[match.kind], Word.InsertLocation.replace);
            });
            await context.sync();
            count++;
        }
        catch (e) {
            console.warn(`Couldn't redact "${match.value}":`, e);
            // Keep going with next match
        }
    }
    return count;
}
async function runRedaction() {
    updateStatus("Analyzing document...");
    const btn = getButton();
    if (btn) {
        btn.disabled = true;
    }
    try {
        await Word.run(async (context) => {
            // Enable tracking if possible
            updateStatus("Enabling track changes...");
            const trackingNote = await setupTrackChanges(context);
            // Add the header
            updateStatus("Adding confidentiality header...");
            await insertHeader(context);
            // Get document text
            updateStatus("Scanning for sensitive information...");
            const body = context.document.body.getRange();
            body.load("text");
            await context.sync();
            // Find all sensitive data
            const matches = findSensitivePatterns(body.text);
            if (matches.length === 0) {
                updateStatus("No sensitive patterns found.");
                return;
            }
            // Redact everything
            const stats = getStats(matches);
            updateStatus(`Found ${stats.total} item${stats.total !== 1 ? "s" : ""} to redact. Processing...`);
            await performRedaction(context, matches);
            // Final sync
            await context.sync();
            const finalMsg = buildMessage(stats, trackingNote);
            updateStatus(finalMsg);
        });
    }
    catch (error) {
        reportError(error);
    }
    finally {
        const btn = getButton();
        if (btn) {
            btn.disabled = false;
        }
    }
}
//# sourceMappingURL=taskpane.js.map