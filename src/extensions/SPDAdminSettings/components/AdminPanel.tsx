import * as React from "react";
import { useState, useEffect, useRef, useCallback } from "react";
import * as ReactDOM from "react-dom";
import styles from "./AdminPanel.module.scss";
import { Dropdown, Icon, Link, Panel, Slider, Toggle } from "@fluentui/react";
import { useBoolean } from "@fluentui/react-hooks";
import { sp } from "@pnp/sp/presets/all";
import {
  IAdminPanelProps,
  ICustomCSSProperties,
  IExtendedUserCustomActionInfo,
} from "./IAdminPanelProps";
import "@pnp/sp/user-custom-actions";
import {
  applyAllCustomCss,
  parseProperties,
  setFooterAlignment,
} from "../utils/helper";

export default function AdminPanel({ context, topString }: IAdminPanelProps) {
  const adminIconRef = useRef<HTMLDivElement | null>(null);
  const [isOpen, { setTrue: openPanel, setFalse: dismissPanel }] =
    useBoolean(false);
  const [fullWidth, { setTrue: setFullWidth, setFalse: unsetFullWidth }] =
    useBoolean(false);
  const [spacing, setSpacing] = useState<number>(null);
  const [alignment, setAlignment] = useState<string>("center");
  const [hasReachedBottom, setHasReachedBottom] = useState(false);

  const websiteURL = context.pageContext.site.absoluteUrl;
  const fromUrl = context.pageContext.legacyPageContext.webTitle
    .toLowerCase()
    .replace(/ & |'| /g, "-");

  useEffect(() => {
    moveIconToTop();
    handleUpdateCustomAction();
    setHasReachedBottom(false);
  }, [context, topString]);

  const handleBottomReached = useCallback(() => {
    setTimeout(() => {
      setFooterAlignment(alignment);
    }, 1000);
  }, []);

  useEffect(() => {
    const scrollContainer = document.querySelector(
      '[data-automation-id="contentScrollRegion"]'
    );
    const checkScrollPosition = () => {
      if (scrollContainer && !hasReachedBottom) {
        const { scrollTop, scrollHeight, clientHeight } = scrollContainer;

        if (scrollTop + clientHeight >= scrollHeight - 5) {
          setHasReachedBottom(true);
          handleBottomReached();
        }
      }
    };
    if (scrollContainer) {
      scrollContainer.addEventListener("scroll", checkScrollPosition, {
        passive: true,
      });
    }
    return () => {
      if (scrollContainer) {
        scrollContainer.removeEventListener("scroll", checkScrollPosition);
      }
    };
  }, [hasReachedBottom, handleBottomReached, topString]);
  const moveIconToTop = () => {
    const headerSection = document.getElementById("HeaderButtonRegion");
    if (headerSection && adminIconRef.current) {
      headerSection.prepend(adminIconRef.current);
    } else {
      setTimeout(moveIconToTop, 2000);
    }
  };

  const handleUpdateCustomAction = async (
    properties?: ICustomCSSProperties
  ) => {
    try {
      const customActions: IExtendedUserCustomActionInfo[] =
        await sp.web.userCustomActions.filter(
          "Location eq 'ClientSideExtension.ApplicationCustomizer'"
        )();

      const customAction = customActions.find(
        (action: any) =>
          action.ClientSideComponentId ===
          "708a1a6a-5d60-455d-b325-0defb4a30f3d"
      );

      if (!customAction) throw new Error("Custom Action not found.");

      const existingProps = parseProperties(
        customAction.ClientSideComponentProperties
      );
      const updatedProps = { ...existingProps, ...properties };

      if (!properties) {
        setDefaultValues(updatedProps);
        applyAllCustomCss(updatedProps);
        return;
      }

      await sp.web.userCustomActions.getById(customAction.Id).update({
        ClientSideComponentProperties: JSON.stringify(updatedProps),
      });
      console.log("Custom action updated successfully!");
    } catch (error) {
      console.error("Error updating custom action:", error);
    }
  };
  /**
   * Parses the custom CSS properties from the ClientSideComponentProperties string.
   * @param properties The ClientSideComponentProperties string to parse.
   * @returns The parsed custom CSS properties.
   */

  /**
   * Sets the default values for the custom CSS properties.
   * @param properties - The custom CSS properties to be set.
   */
  const setDefaultValues = (properties: ICustomCSSProperties) => {
    properties.fullWidth ? setFullWidth() : unsetFullWidth();
    setSpacing(properties.spacing ?? 24);
    setAlignment(properties.alignment ?? "center");
  };

  /**
   * Handles changes to custom CSS properties and updates the settings.
   * @param value - The new value for the property being changed.
   * @param field - The specific field of ICustomCSSProperties being updated.
   */
  const handleChange = (value: any, field: keyof ICustomCSSProperties) => {
    const updatedSettings: ICustomCSSProperties = {
      alignment: field === "alignment" ? value : alignment,
      fullWidth: field === "fullWidth" ? value : fullWidth,
      spacing: field === "spacing" ? value : spacing,
    };

    setDefaultValues(updatedSettings);
    applyAllCustomCss(updatedSettings);
    handleUpdateCustomAction(updatedSettings);
  };

  return ReactDOM.createPortal(
    <div className={styles.app} ref={adminIconRef}>
      <div className={styles.top} id="adminIconClick">
        <div className={styles.iconContainer} onClick={openPanel}>
          <Icon iconName="ContentSettings" className={styles.defaultIcon} />
        </div>
      </div>

      <Panel
        headerText="Admin Settings"
        isOpen={isOpen}
        onDismiss={dismissPanel}
      >
        <div className={styles.container}>
          <Link
            href={`${websiteURL}/SitePages/Admin-Page.aspx#${fromUrl}`}
            className={styles.items}
            target="_blank"
            data-interception="off"
          >
            <Icon iconName="ContentSettings" className={styles.icons} /> Admin
            Page
          </Link>

          <Toggle
            label="Enable full width for all screen size"
            checked={fullWidth}
            onText="On"
            offText="Off"
            onChange={(_, checked) => handleChange(!!checked, "fullWidth")}
          />

          <Slider
            label="Webparts Spacing"
            min={0}
            max={30}
            step={1}
            value={spacing}
            onChange={(value) => handleChange(value, "spacing")}
          />

          <Dropdown
            label="Footer alignment"
            selectedKey={alignment}
            options={[
              { key: "left", text: "Left" },
              { key: "center", text: "Center" },
              { key: "right", text: "Right" },
            ]}
            onChange={(_, option) => {
              handleChange(option?.key || "center", "alignment");
            }}
          />
        </div>
      </Panel>
    </div>,
    document.body
  );
}
