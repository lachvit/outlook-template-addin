import { makeStyles } from "@fluentui/react-components";

const useStyles = makeStyles({
  frame: {
    flex: "1 1 auto",
    width: "100%",
    border: "none",
    backgroundColor: "#ffffff",
  },
});

interface TemplatePreviewProps {
  /** Must already be sanitized. */
  html: string;
  title: string;
}

export function TemplatePreview({ html, title }: TemplatePreviewProps) {
  const styles = useStyles();
  const srcDoc =
    "<!DOCTYPE html><html><head><meta charset=\"utf-8\">" +
    "<meta http-equiv=\"Content-Security-Policy\" content=\"script-src 'none'\">" +
    "<style>body{margin:8px;font-family:Calibri,Arial,sans-serif;font-size:11pt;}img{max-width:100%;height:auto;}</style>" +
    `</head><body>${html}</body></html>`;

  // An empty sandbox blocks scripts, forms, popups and same-origin access.
  return <iframe className={styles.frame} title={`Preview of ${title}`} sandbox="" srcDoc={srcDoc} />;
}
