import * as React from "react";
import {
  Body1,
  Body1Strong,
  Button,
  Caption1,
  Dropdown,
  Field,
  MessageBar,
  MessageBarActions,
  MessageBarBody,
  MessageBarTitle,
  Option,
  SearchBox,
  Spinner,
  Toast,
  ToastBody,
  ToastTitle,
  Toaster,
  makeStyles,
  mergeClasses,
  tokens,
  useId,
  useToastController,
} from "@fluentui/react-components";
import { getGraphToken, signIn, SignInRequiredError } from "../lib/auth";
import { getMissingSettings, SHAREPOINT_FOLDER, SHAREPOINT_LIBRARY, SHAREPOINT_SITE_URL } from "../lib/config";
import { filterTemplates, getCategories } from "../lib/filter";
import { insertIntoMessage } from "../lib/outlook";
import { SharePointTemplateStore } from "../lib/sharePoint";
import { PreparedTemplate, prepareTemplate } from "../lib/templateHtml";
import { TemplateInfo } from "../lib/templateItems";
import { TemplatePreview } from "./TemplatePreview";

type IndexState =
  | { status: "loading" }
  | { status: "signin"; message?: string }
  | { status: "error"; message: string }
  | { status: "ready"; templates: TemplateInfo[] };

type PreviewState =
  | { status: "idle" }
  | { status: "loading" }
  | { status: "error"; message: string }
  | { status: "ready"; template: PreparedTemplate };

const ALL_CATEGORIES = "";

const missingSettings = getMissingSettings();
const store = new SharePointTemplateStore(
  { siteUrl: SHAREPOINT_SITE_URL, libraryName: SHAREPOINT_LIBRARY, folderPath: SHAREPOINT_FOLDER },
  getGraphToken,
);

function errorMessage(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}

const useStyles = makeStyles({
  root: {
    display: "flex",
    flexDirection: "column",
    gap: tokens.spacingVerticalM,
    height: "100vh",
    boxSizing: "border-box",
    padding: tokens.spacingHorizontalM,
  },
  list: {
    flex: "1 1 40%",
    minHeight: "120px",
    overflowY: "auto",
    display: "flex",
    flexDirection: "column",
    gap: tokens.spacingVerticalXXS,
    border: `${tokens.strokeWidthThin} solid ${tokens.colorNeutralStroke2}`,
    borderRadius: tokens.borderRadiusMedium,
    padding: tokens.spacingVerticalXS,
  },
  item: {
    display: "flex",
    flexDirection: "column",
    alignItems: "flex-start",
    gap: "2px",
    width: "100%",
    textAlign: "left",
    padding: `${tokens.spacingVerticalS} ${tokens.spacingHorizontalM}`,
    border: "none",
    borderRadius: tokens.borderRadiusMedium,
    backgroundColor: "transparent",
    color: tokens.colorNeutralForeground1,
    fontFamily: tokens.fontFamilyBase,
    cursor: "pointer",
    ":hover": { backgroundColor: tokens.colorNeutralBackground1Hover },
    ":focus-visible": { outline: `2px solid ${tokens.colorStrokeFocus2}` },
  },
  selected: {
    backgroundColor: tokens.colorBrandBackground2,
    ":hover": { backgroundColor: tokens.colorBrandBackground2Hover },
  },
  preview: {
    flex: "1 1 60%",
    minHeight: "180px",
    display: "flex",
    flexDirection: "column",
    border: `${tokens.strokeWidthThin} solid ${tokens.colorNeutralStroke2}`,
    borderRadius: tokens.borderRadiusMedium,
    overflow: "hidden",
  },
  centered: {
    flex: "1 1 auto",
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    textAlign: "center",
    padding: tokens.spacingHorizontalL,
  },
  muted: { color: tokens.colorNeutralForeground3 },
});

export function App() {
  const styles = useStyles();
  const toasterId = useId("toaster");
  const { dispatchToast } = useToastController(toasterId);

  const [index, setIndex] = React.useState<IndexState>({ status: "loading" });
  const [query, setQuery] = React.useState("");
  const [category, setCategory] = React.useState(ALL_CATEGORIES);
  const [selectedId, setSelectedId] = React.useState<string>();
  const [preview, setPreview] = React.useState<PreviewState>({ status: "idle" });
  const [inserting, setInserting] = React.useState(false);
  const [signingIn, setSigningIn] = React.useState(false);
  const prepared = React.useRef(new Map<string, PreparedTemplate>());

  const loadIndex = React.useCallback(() => {
    if (missingSettings.length > 0) {
      setIndex({
        status: "error",
        message: `The add-in was built without these settings: ${missingSettings.join(", ")}. Ask IT to finish the setup.`,
      });
      return;
    }
    setIndex({ status: "loading" });
    prepared.current.clear();
    store.listTemplates().then(
      (templates) => setIndex({ status: "ready", templates }),
      (error: unknown) =>
        setIndex(
          error instanceof SignInRequiredError ? { status: "signin" } : { status: "error", message: errorMessage(error) },
        ),
    );
  }, []);

  const startSignIn = async () => {
    setSigningIn(true);
    try {
      await signIn();
      loadIndex();
    } catch (error) {
      setIndex({ status: "signin", message: errorMessage(error) });
    } finally {
      setSigningIn(false);
    }
  };

  React.useEffect(() => loadIndex(), [loadIndex]);

  const templates = React.useMemo(() => (index.status === "ready" ? index.templates : []), [index]);
  const categories = React.useMemo(() => getCategories(templates), [templates]);
  const visible = React.useMemo(() => filterTemplates(templates, query, category), [templates, query, category]);
  const selected = React.useMemo(() => templates.find((t) => t.id === selectedId), [templates, selectedId]);

  React.useEffect(() => {
    if (!selected) {
      setPreview({ status: "idle" });
      return;
    }
    const cached = prepared.current.get(selected.id);
    if (cached) {
      setPreview({ status: "ready", template: cached });
      return;
    }
    let cancelled = false;
    setPreview({ status: "loading" });
    store
      .getTemplateHtml(selected.id)
      .then((raw) => {
        const template = prepareTemplate(raw);
        prepared.current.set(selected.id, template);
        if (!cancelled) {
          setPreview({ status: "ready", template });
        }
      })
      .catch((error: unknown) => {
        if (!cancelled) {
          setPreview({ status: "error", message: errorMessage(error) });
        }
      });
    return () => {
      cancelled = true;
    };
  }, [selected]);

  const insert = async () => {
    if (preview.status !== "ready") {
      return;
    }
    setInserting(true);
    try {
      const format = await insertIntoMessage(preview.template.html, preview.template.text);
      if (format === "html") {
        dispatchToast(
          <Toast>
            <ToastTitle>Template inserted</ToastTitle>
          </Toast>,
          { intent: "success" },
        );
      } else {
        dispatchToast(
          <Toast>
            <ToastTitle>Inserted as plain text</ToastTitle>
            <ToastBody>This message is in plain-text format, so formatting and images were left out.</ToastBody>
          </Toast>,
          { intent: "warning" },
        );
      }
    } catch (error) {
      dispatchToast(
        <Toast>
          <ToastTitle>Could not insert the template</ToastTitle>
          <ToastBody>{errorMessage(error)}</ToastBody>
        </Toast>,
        { intent: "error" },
      );
    } finally {
      setInserting(false);
    }
  };

  if (index.status === "loading") {
    return (
      <div className={styles.root}>
        <div className={styles.centered}>
          <Spinner label="Loading templates..." />
        </div>
      </div>
    );
  }

  if (index.status === "signin") {
    return (
      <div className={styles.root}>
        <MessageBar intent={index.message ? "warning" : "info"} layout="multiline">
          <MessageBarBody>
            <MessageBarTitle>Sign in to see the templates</MessageBarTitle>
            {index.message ?? "The templates are stored in SharePoint. Sign in with your work account to load them."}
          </MessageBarBody>
        </MessageBar>
        <Button appearance="primary" disabled={signingIn} onClick={() => void startSignIn()}>
          {signingIn ? "Signing in..." : "Sign in"}
        </Button>
      </div>
    );
  }

  if (index.status === "error") {
    return (
      <div className={styles.root}>
        <MessageBar intent="error" layout="multiline">
          <MessageBarBody>
            <MessageBarTitle>Templates are unavailable</MessageBarTitle>
            {index.message}
          </MessageBarBody>
          <MessageBarActions>
            <Button onClick={loadIndex}>Try again</Button>
          </MessageBarActions>
        </MessageBar>
      </div>
    );
  }

  return (
    <div className={styles.root}>
      <Toaster toasterId={toasterId} position="bottom" />

      <Field label="Search">
        <SearchBox
          placeholder="Name, description or keyword"
          value={query}
          onChange={(_, data) => setQuery(data.value)}
        />
      </Field>

      <Field label="Category">
        <Dropdown
          value={category || "All categories"}
          selectedOptions={[category]}
          onOptionSelect={(_, data) => setCategory(data.optionValue ?? ALL_CATEGORIES)}
        >
          <Option value={ALL_CATEGORIES}>All categories</Option>
          {categories.map((c) => (
            <Option key={c} value={c}>
              {c}
            </Option>
          ))}
        </Dropdown>
      </Field>

      <div className={styles.list} role="list" aria-label="Templates">
        {visible.length === 0 ? (
          <div className={styles.centered}>
            <Body1 className={styles.muted}>
              {templates.length === 0 ? "No templates have been published yet." : "No templates match your search."}
            </Body1>
          </div>
        ) : (
          visible.map((t) => (
            <div role="listitem" key={t.id}>
              <button
                type="button"
                className={mergeClasses(styles.item, t.id === selectedId && styles.selected)}
                aria-pressed={t.id === selectedId}
                onClick={() => setSelectedId(t.id)}
              >
                <Body1Strong>{t.name}</Body1Strong>
                <Caption1 className={styles.muted}>
                  {t.description ? `${t.category} - ${t.description}` : t.category}
                </Caption1>
              </button>
            </div>
          ))
        )}
      </div>

      <section className={styles.preview} aria-label="Preview">
        {preview.status === "idle" && (
          <div className={styles.centered}>
            <Body1 className={styles.muted}>Select a template to see a preview.</Body1>
          </div>
        )}
        {preview.status === "loading" && (
          <div className={styles.centered}>
            <Spinner size="small" label="Loading preview..." />
          </div>
        )}
        {preview.status === "error" && (
          <MessageBar intent="error">
            <MessageBarBody>{preview.message}</MessageBarBody>
          </MessageBar>
        )}
        {preview.status === "ready" && <TemplatePreview html={preview.template.html} title={selected?.name ?? ""} />}
      </section>

      <Button appearance="primary" disabled={preview.status !== "ready" || inserting} onClick={() => void insert()}>
        {inserting ? "Inserting..." : "Insert template"}
      </Button>
    </div>
  );
}
