import { insertIntoMessage } from "../src/taskpane/lib/outlook";

type Callback<T> = (result: { status: string; value?: T; error?: { message: string } }) => void;

/** Minimal fake of Office.js: a body with a cursor that may or may not be inside it. */
function setupOffice(options: { html: boolean; cursorInBody: boolean; body?: string }) {
  let body = options.body ?? "<p>Signature</p>";
  const ok = <T,>(cb: Callback<T>, value?: T) => cb({ status: "succeeded", value });
  const calls: string[] = [];
  const item = {
    body: {
      getTypeAsync: (cb: Callback<string>) => ok(cb, options.html ? "html" : "text"),
      getAsync: (_type: string, cb: Callback<string>) => ok(cb, body),
      setSelectedDataAsync: (data: string, _opts: unknown, cb: Callback<void>) => {
        calls.push("setSelectedData");
        if (options.cursorInBody) {
          body = body + data;
        }
        ok(cb);
      },
      prependAsync: (data: string, _opts: unknown, cb: Callback<void>) => {
        calls.push("prepend");
        body = data + body;
        ok(cb);
      },
    },
  };
  (globalThis as unknown as { Office: unknown }).Office = {
    context: { mailbox: { item } },
    AsyncResultStatus: { Succeeded: "succeeded", Failed: "failed" },
    CoercionType: { Html: "html", Text: "text" },
  };
  return { calls, body: () => body };
}

describe("insertIntoMessage", () => {
  it("inserts at the cursor when the cursor is in the body", async () => {
    const office = setupOffice({ html: true, cursorInBody: true });
    await expect(insertIntoMessage("<p>T</p>", "T")).resolves.toEqual({ format: "html", position: "cursor" });
    expect(office.calls).toEqual(["setSelectedData"]);
    expect(office.body()).toBe("<p>Signature</p><p>T</p>");
  });

  it("adds the template at the top of the body when the cursor is in To, Cc or Subject", async () => {
    const office = setupOffice({ html: true, cursorInBody: false });
    await expect(insertIntoMessage("<p>T</p>", "T")).resolves.toEqual({ format: "html", position: "top" });
    expect(office.calls).toEqual(["setSelectedData", "prepend"]);
    expect(office.body()).toBe("<p>T</p><p>Signature</p>");
  });

  it("uses the text version for plain-text messages", async () => {
    const office = setupOffice({ html: false, cursorInBody: false, body: "Sig" });
    await expect(insertIntoMessage("<p>T</p>", "T")).resolves.toEqual({ format: "text", position: "top" });
    expect(office.body()).toBe("TSig");
  });

  it("explains when no message is being written", async () => {
    (globalThis as unknown as { Office: unknown }).Office = { context: { mailbox: { item: undefined } } };
    await expect(insertIntoMessage("<p>T</p>", "T")).rejects.toThrow(/Open a message you are writing/);
  });
});
