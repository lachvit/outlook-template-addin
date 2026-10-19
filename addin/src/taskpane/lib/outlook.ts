export type InsertedFormat = "html" | "text";

/** Where the template ended up: at the cursor in the body, or at the top of the body. */
export type InsertedPosition = "cursor" | "top";

export interface InsertResult {
  format: InsertedFormat;
  position: InsertedPosition;
}

function currentComposeItem(): Office.MessageCompose {
  const item = Office.context.mailbox?.item as unknown as Office.MessageCompose | undefined;
  if (!item || !item.body) {
    throw new Error("Open a message you are writing, then try again.");
  }
  return item;
}

function call<T>(start: (callback: (result: Office.AsyncResult<T>) => void) => void): Promise<T> {
  return new Promise((resolve, reject) => {
    start((result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve(result.value);
      } else {
        reject(new Error(result.error.message));
      }
    });
  });
}

/**
 * Inserts into the message body, never into To, Cc or Subject.
 * The template goes at the cursor (or replaces the selection) when the cursor is in the body. When it is
 * elsewhere, Outlook does not change the body, so the body is compared before and after and the template
 * is then added at the top of the body (above the signature and any quoted reply).
 * Plain-text messages get the text version.
 */
export async function insertIntoMessage(html: string, text: string): Promise<InsertResult> {
  const item = currentComposeItem();
  const bodyType = await call<Office.CoercionType>((cb) => item.body.getTypeAsync(cb));
  const format: InsertedFormat = bodyType === Office.CoercionType.Html ? "html" : "text";
  const coercionType = format === "html" ? Office.CoercionType.Html : Office.CoercionType.Text;
  const data = format === "html" ? html : text;

  const before = await call<string>((cb) => item.body.getAsync(coercionType, cb));
  let insertError: unknown;
  try {
    await call<void>((cb) => item.body.setSelectedDataAsync(data, { coercionType }, cb));
    const after = await call<string>((cb) => item.body.getAsync(coercionType, cb));
    if (after !== before) {
      return { format, position: "cursor" };
    }
  } catch (error) {
    // Some Outlook versions reject the call ("The operation is not supported") when the cursor isn't in the body.
    insertError = error;
  }

  try {
    await call<void>((cb) => item.body.prependAsync(data, { coercionType }, cb));
  } catch (error) {
    const reason = error instanceof Error ? error.message : String(error);
    const first = insertError instanceof Error && insertError.message !== reason ? ` (${insertError.message})` : "";
    throw new Error(`Outlook could not add the template to the message: ${reason}${first}`);
  }
  return { format, position: "top" };
}
