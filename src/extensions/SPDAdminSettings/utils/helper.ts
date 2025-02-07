import { ICustomCSSProperties } from "../components/IAdminPanelProps";

export const parseProperties = (
  properties: string | undefined
): ICustomCSSProperties => {
  return properties ? JSON.parse(properties) : {};
};

export const applyAllCustomCss = (properties: ICustomCSSProperties) => {
  setWidth(properties.fullWidth);
  setSpacing(properties.spacing);
  setFooterAlignment(properties.alignment);
};

export const setWidth = (fullWidth: boolean) => {
  document
    .querySelectorAll('[data-automation-id*="CanvasZone-SectionContainer"]')
    .forEach((section: HTMLElement) => {
      section.style.maxWidth = fullWidth ? "100%" : "revert-layer";
    });
};
export const setSpacing = (spacing: number) => {
  document
    .querySelectorAll('[data-automation-id="CanvasControl"]')
    .forEach((container: HTMLElement) => {
      container.style.margin = `${spacing ?? 24}px 0`;
    });
};
export const setFooterAlignment = (alignment: string) => {
  document
    .querySelectorAll('[data-automationid="SimpleFooter"]')
    .forEach((section: HTMLElement) => {
      section.style.justifyContent = alignment || "center";
    });
};
