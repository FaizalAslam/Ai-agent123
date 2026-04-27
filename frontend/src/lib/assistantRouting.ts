export type AssistantRoute =
  | {
      kind: "office";
      endpoint: "/office/execute";
      app: "excel" | "word" | "powerpoint";
      raw: string;
      command: string;
      reason: string;
    }
  | {
      kind: "system";
      endpoint: "/execute";
      command: string;
      reason: string;
    }
  | {
      kind: "unsupported";
      command: string;
      reason: string;
    };

const officeTargets = {
  excel: ["excel", "spreadsheet", "workbook", "worksheet", "sheet", "xlsx", "xlsm", "xls", "csv"],
  word: ["word", "document", "docx", "doc"],
  powerpoint: ["powerpoint", "power point", "ppt", "pptx", "presentation", "slide deck", "slides", "slide", "deck"]
} as const;

const actionTerms = [
  "create",
  "make",
  "generate",
  "build",
  "new",
  "open",
  "save as",
  "add",
  "insert",
  "edit",
  "update",
  "modify",
  "write",
  "format",
  "table",
  "chart",
  "row",
  "column",
  "cell",
  "slide",
  "paragraph",
  "heading",
  "bold",
  "italic",
  "underline",
  "color",
  "background",
  "border",
  "formula",
  "bullet",
  "title",
  "save",
  "close",
  "deck"
];

const launchPrefixes = ["open ", "launch ", "start ", "run ", "boot "];

function containsTerm(text: string, term: string) {
  if (term.includes(" ")) return text.includes(term);
  return new RegExp(`\\b${term.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")}\\b`).test(text);
}

export function classifyAssistantCommand(command: string): AssistantRoute {
  const original = command.trim();
  const text = original.toLowerCase().replace(/\s+/g, " ");

  const prefixMatch = original.match(/^agent\s*:\s*(excel|word|powerpoint|ppt)\s*:\s*(.+)$/i);
  if (prefixMatch) {
    const app = prefixMatch[1].toLowerCase() === "ppt" ? "powerpoint" : (prefixMatch[1].toLowerCase() as "excel" | "word" | "powerpoint");
    return {
      kind: "office",
      endpoint: "/office/execute",
      app,
      raw: prefixMatch[2].trim(),
      command: original,
      reason: "agent prefix"
    };
  }

  if (/^close\s+(?:the\s+)?(?:current\s+)?(?:document|file)\b/.test(text)) {
    return {
      kind: "office",
      endpoint: "/office/execute",
      app: "word",
      raw: original,
      command: original,
      reason: "document lifecycle command"
    };
  }

  if (
    text.startsWith("close ") &&
    /\b(excel|word|powerpoint|power point|ppt|microsoft word)\b/.test(text) &&
    !/\b(file|workbook|worksheet|spreadsheet|document|docx|xlsx|pptx|presentation|slides|deck)\b/.test(text)
  ) {
    return {
      kind: "system",
      endpoint: "/execute",
      command: original,
      reason: "app-close command"
    };
  }

  const hasAction = actionTerms.some((term) => containsTerm(text, term));
  for (const [app, targets] of Object.entries(officeTargets)) {
    if (targets.some((target) => containsTerm(text, target)) && hasAction) {
      return {
        kind: "office",
        endpoint: "/office/execute",
        app: app as "excel" | "word" | "powerpoint",
        raw: original,
        command: original,
        reason: "office keyword with document action"
      };
    }
  }

  if (launchPrefixes.some((prefix) => text.startsWith(prefix))) {
    return {
      kind: "system",
      endpoint: "/execute",
      command: original,
      reason: "app-launch command"
    };
  }

  return {
    kind: "unsupported",
    command: original,
    reason: "Not recognized as an Office automation or app-launch command."
  };
}
