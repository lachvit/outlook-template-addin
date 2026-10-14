export type InsertedFormat = "html" | "text";

function currentComposeItem(): Office.MessageCompose {
  const item = Office.context.mailbox?.item as unknown as Office.MessageCompose | undefined;
  if (!item || !item.body) {
    throw new Error("Open a message you are writing, then try again.");
  }
  return item;
}

function getBodyType(item: Office.MessageCompose): Promise<Office.CoercionType> {
  return new Promise((resolve, reject) => {
    item.body.getTypeAsync((result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve(result.value);
      } else {
        reject(new Error(result.error.message));
      }
    });
  });
}

function setSelectedData(item: Office.MessageCompose, data: string, coercionType: Office.CoercionType): Promise<void> {
  return new Promise((resolve, reject) => {
    item.body.setSelectedDataAsync(data, { coercionType }, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve();
      } else {
        reject(new Error(result.error.message));
      }
    });
  });
}

/** Inserts at the cursor (or replaces the selection). Plain-text messages get the text version. */
export async function insertIntoMessage(html: string, text: string): Promise<InsertedFormat> {
  const item = currentComposeItem();
  if ((await getBodyType(item)) === Office.CoercionType.Html) {
    await setSelectedData(item, html, Office.CoercionType.Html);
    return "html";
  }
  await setSelectedData(item, text, Office.CoercionType.Text);
  return "text";
}
