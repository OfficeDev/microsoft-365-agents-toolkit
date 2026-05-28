// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import "./sampleFilter.scss";

import { debounce } from "lodash";
import * as React from "react";

import { ActionButton, Dropdown, IDropdownOption, IDropdownStyles, IStyle } from "@fluentui/react";
import { VSCodeButton, VSCodeTextField } from "@vscode/webview-ui-toolkit/react";

import {
  TelemetryEvent,
  TelemetryProperty,
  TelemetryTriggerFrom,
} from "../../telemetry/extTelemetryEvents";
import { Commands } from "../Commands";
import { Grid } from "../resources";
import { SampleFilterProps } from "./ISamples";

export default class SampleFilter extends React.Component<SampleFilterProps, unknown> {
  constructor(props: SampleFilterProps) {
    super(props);
  }

  render() {
    const sampleTypes = this.props.filterOptions.capabilities;
    const sampleLanguages = this.props.filterOptions.languages;
    const sampleTechniques = this.props.filterOptions.technologies;
    const typeOptions: IDropdownOption[] = sampleTypes.map((type) => {
      const selected = this.props.filterTags.indexOf(type) >= 0;
      return { key: type, text: type, selected };
    });
    const languageOptions: IDropdownOption[] = sampleLanguages.map((type) => {
      const selected = this.props.filterTags.indexOf(type) >= 0;
      return { key: type, text: type, selected };
    });
    const techniqueOptions: IDropdownOption[] = sampleTechniques.map((type) => {
      const selected = this.props.filterTags.indexOf(type) >= 0;
      return { key: type, text: type, selected };
    });
    const dropdownStyles = this.getDropdownStyles();
    return (
      <div className="sample-filter">
        <div className="sample-filter-bar">
          <VSCodeTextField
            className="search-box"
            placeholder="Search samples"
            value={this.props.query}
            onInput={this.onSearchTextChanged}
          >
            <span slot="start" className="codicon codicon-search"></span>
          </VSCodeTextField>
          <Dropdown
            ariaLabel="Select to filter platform capability:"
            placeholder="Platform capability"
            multiSelect
            options={typeOptions}
            styles={dropdownStyles}
            onChange={this.onFilterTagsChanged}
            selectedKeys={sampleTypes.filter((type) => {
              return this.props.filterTags.indexOf(type) >= 0;
            })}
            dropdownWidth="auto"
          />
          <Dropdown
            ariaLabel="Select to filter programming language:"
            placeholder="Language"
            multiSelect
            options={languageOptions}
            styles={dropdownStyles}
            onChange={this.onFilterTagsChanged}
            selectedKeys={sampleLanguages.filter((type) => {
              return this.props.filterTags.indexOf(type) >= 0;
            })}
            dropdownWidth="auto"
          />
          <Dropdown
            ariaLabel="Select to filter technology:"
            placeholder="Technology"
            multiSelect
            options={techniqueOptions}
            styles={dropdownStyles}
            onChange={this.onFilterTagsChanged}
            selectedKeys={sampleTechniques.filter((type) => {
              return this.props.filterTags.indexOf(type) >= 0;
            })}
            dropdownWidth="auto"
          />
          <div className="filter-bar"></div>
          <VSCodeButton
            onClick={() => this.props.onLayoutChanged("grid")}
            appearance="icon"
            aria-label="gallary view"
            className={`layout-button ${this.props.layout === "grid" ? "layout-selected" : ""}`}
          >
            <Grid />
          </VSCodeButton>
          <VSCodeButton
            onClick={() => this.props.onLayoutChanged("list")}
            appearance="icon"
            aria-label="list view"
            className={`layout-button ${this.props.layout === "list" ? "layout-selected" : ""}`}
          >
            <span className="codicon codicon-list-unordered"></span>
          </VSCodeButton>
        </div>
        <div className="filter-tag-bar">
          {this.props.filterTags.map((tag) => (
            <div className="filter-tag">
              <span>{tag}</span>
              <span className="codicon codicon-close" onClick={() => this.onTagRemoved(tag)}></span>
            </div>
          ))}
          {this.props.filterTags.length > 0 && (
            <ActionButton onClick={this.onAllTagsRemoved}>Clear all</ActionButton>
          )}
        </div>
      </div>
    );
  }

  private onSearchTextChanged = (e: { target: { value: string } }) => {
    debounce(() => {
      vscode.postMessage({
        command: Commands.SendTelemetryEvent,
        data: {
          eventName: TelemetryEvent.SearchSample,
          properties: {
            [TelemetryProperty.TriggerFrom]: TelemetryTriggerFrom.SampleGallery,
            [TelemetryProperty.SearchText]: e.target.value,
            [TelemetryProperty.SampleFilters]: this.props.filterTags.join(","),
          },
        },
      });
      this.props.onFilterConditionChanged(e.target.value, this.props.filterTags);
    }, 500)();
  };

  private onFilterTagChanged = (
    telemetryEvent: TelemetryEvent,
    changedFilter: string,
    newFilterTags: string[]
  ) => {
    vscode.postMessage({
      command: Commands.SendTelemetryEvent,
      data: {
        eventName: telemetryEvent,
        properties: {
          [TelemetryProperty.TriggerFrom]: TelemetryTriggerFrom.SampleGallery,
          [TelemetryProperty.ChangedFilter]: changedFilter,
          [TelemetryProperty.SampleFilters]: this.props.filterTags.join(","),
        },
      },
    });
    this.props.onFilterConditionChanged(this.props.query, newFilterTags);
  };

  private onFilterTagsChanged = (
    _event: React.FormEvent<HTMLDivElement>,
    option?: IDropdownOption
  ) => {
    const choice = option?.key as string;
    let telemetryEvent = TelemetryEvent.FilterSampleAdd;
    let newData: string[] = [];
    if (option?.selected) {
      newData = [...this.props.filterTags, choice];
    } else {
      telemetryEvent = TelemetryEvent.FilterSampleRemove;
      newData = this.props.filterTags.filter((tag) => tag !== choice);
    }
    this.onFilterTagChanged(telemetryEvent, choice, newData);
  };

  private onTagRemoved = (removedTag: string) => {
    vscode.postMessage({
      command: Commands.SendTelemetryEvent,
      data: {
        eventName: TelemetryEvent.FilterSampleRemove,
        properties: {
          [TelemetryProperty.TriggerFrom]: TelemetryTriggerFrom.SampleGallery,
          [TelemetryProperty.ChangedFilter]: removedTag,
          [TelemetryProperty.SampleFilters]: this.props.filterTags.join(","),
        },
      },
    });
    const newFilterTags = this.props.filterTags.filter((tag) => tag !== removedTag);
    this.props.onFilterConditionChanged(this.props.query, newFilterTags);
  };

  private onAllTagsRemoved = () => {
    vscode.postMessage({
      command: Commands.SendTelemetryEvent,
      data: {
        eventName: TelemetryEvent.FilterSampleRemove,
        properties: {
          [TelemetryProperty.TriggerFrom]: TelemetryTriggerFrom.SampleGallery,
          [TelemetryProperty.ChangedFilter]: this.props.filterTags.join(","),
          [TelemetryProperty.SampleFilters]: this.props.filterTags.join(","),
        },
      },
    });
    this.props.onFilterConditionChanged(this.props.query, []);
  };

  private getDropdownStyles = (): Partial<IDropdownStyles> => {
    const dropDownStyle: IStyle = {
      "span:first-child": {
        height: 24,
        lineHeight: 21,
        backgroundColor: "rgba(255, 255, 255, 0.1)",
        color: "var(--vscode-peekViewTitleDescription-foreground, #CCCCCC)",
        fontSize: 13,
        border: "1px solid var(--vscode-menu-separatorBackground, #3C3C3C)",
        fontFamily: "var(--font-family)",
        width: 152,
      },
    };
    const caretStyle: IStyle = {
      color: "var(--vscode-dropdown-foreground, #CCCCCC)",
      fontSize: 11,
      lineHeight: 16,
    };
    const checkboxStyle: IStyle = {
      ".ms-Checkbox-checkbox": {
        backgroundColor: "var(--vscode-dropdown-background, #3C3C3C)",
        border: "1px solid var(--vscode-button-secondaryHoverBackground, #3C3C3C)",
        i: {
          color: "var(--vscode-dropdown-background, #3C3C3C)",
        },
      },
    };
    const checkboxStyleSelected: IStyle = {
      ".ms-Checkbox-checkbox": {
        backgroundColor: "var(--vscode-dropdown-background, #3C3C3C)",
        border: "1px solid var(--vscode-button-secondaryHoverBackground, #3C3C3C)",
        i: {
          color: "var(--vscode-peekViewTitleDescription-foreground, #cccccc)",
        },
      },
    };
    const dropdownStyles: Partial<IDropdownStyles> = {
      dropdown: {
        ...dropDownStyle,
        ":hover": {
          ...dropDownStyle,
          ".ms-Dropdown-caretDown": {
            color: "var(--vscode-dropdown-foreground, #CCCCCC)",
          },
        },
        ":focus": {
          ...dropDownStyle,
          ".ms-Dropdown-caretDown": {
            color: "var(--vscode-dropdown-foreground, #CCCCCC)",
          },
        },
        ":active": {
          ".ms-Dropdown-caretDown": {
            color: "var(--vscode-dropdown-foreground, #CCCCCC)",
          },
        },
        marginLeft: 16,
      },
      caretDown: {
        ...caretStyle,
      },
      caretDownWrapper: {
        height: 24,
        lineHeight: 24,
        color: "var(--vscode-dropdown-foreground, #CCCCCC)",
      },
      callout: {
        ".ms-Callout-main": {
          border: "1px solid var(--vscode-inputValidation-infoBorder, #007ACC)",
        },
      },
      dropdownItemsWrapper: {
        padding: "4px 0",
        backgroundColor: "var(--vscode-editorGroupHeader-tabsBackground, #252526)",
      },
      dropdownItem: {
        backgroundColor: "var(--vscode-editorGroupHeader-tabsBackground, #252526)",
        minHeight: 22,
        height: 22,
        ...checkboxStyle,
        ":active": {
          backgroundColor: "var(--vscode-editorGroupHeader-tabsBackground, #252526) !important",
        },
        "input:focus + .ms-Checkbox-label": {
          ...checkboxStyle,
        },
        "input:focus + .ms-Checkbox-label .ms-Checkbox-checkbox": {
          borderColor: "var(--vscode-inputValidation-infoBorder, #007ACC)",
        },
        "input:focus + .ms-Checkbox-label .ms-Checkbox-checkmark": {
          color: "var(--vscode-dropdown-background, #3C3C3C)",
        },
        ":hover": {
          backgroundColor: "var(--vscode-editorStickyScrollHover-background, #303031) !important",
          ".ms-Checkbox-checkmark": {
            color: "var(--vscode-dropdown-background, #3C3C3C)",
          },
          ".ms-Checkbox-checkbox": {
            borderColor: "var(--vscode-button-secondaryHoverBackground, #3C3C3C)",
          },
        },
      },
      dropdownItemSelected: {
        minHeight: 22,
        height: 22,
        backgroundColor: "var(--vscode-editorGroupHeader-tabsBackground, #252526)",
        ...checkboxStyleSelected,
        ":active": {
          backgroundColor: "var(--vscode-editorGroupHeader-tabsBackground, #252526) !important",
        },
        "input:focus + .ms-Checkbox-label": {
          ...checkboxStyleSelected,
        },
        "input:focus + .ms-Checkbox-label .ms-Checkbox-checkbox": {
          borderColor: "var(--vscode-inputValidation-infoBorder, #007ACC)",
        },
        ":focus": {
          ...checkboxStyleSelected,
        },
        ":hover": {
          backgroundColor: "var(--vscode-editorStickyScrollHover-background, #303031) !important",
          ...checkboxStyleSelected,
        },
      },
      dropdownOptionText: {
        fontSize: "13px",
        color: "var(--vscode-dropdown-foreground, #CCCCCC)",
      },
    };
    return dropdownStyles;
  };
}
