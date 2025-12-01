import React, { Component } from "react";

import { mergeStyles, mergeStyleSets } from "@fluentui/react";
import { tokens } from "@fluentui/react-components";

/**
 * Style definitions for the widget elements
 * @internal
 */
const classNames = mergeStyleSets({
  root: {
    display: "grid",
    padding: "1.25rem 2rem 1.25rem 2rem",
    backgroundColor: tokens.colorNeutralBackground1,
    border: "1px solid var(--colorTransparentStroke)",
    boxShadow: tokens.shadow4,
    borderRadius: tokens.borderRadiusMedium,
    gap: tokens.spacingHorizontalL,
    gridTemplateRows: "max-content 1fr max-content",
  },
  header: {
    display: "grid",
    height: "max-content",
    "& div": {
      display: "grid",
      gap: tokens.spacingHorizontalS,
      alignItems: "center",
      gridTemplateColumns: "min-content 1fr min-content",
    },
    "& svg": {
      height: "1.5rem",
      width: "1.5rem",
    },
    "& span": {
      fontWeight: tokens.fontWeightSemibold,
      lineHeight: tokens.lineHeightBase200,
      fontSize: tokens.fontSizeBase200,
    },
  },
  footer: {
    "& button": {
      width: "fit-content",
    },
  },
});

/**
 * The base component that provides basic functionality to create a widget.
 */
export class BaseWidget extends Component {
  /**
   * Constructor of BaseWidget.
   * @param {Object} props - The props of the component.
   */
  constructor(props) {
    super(props);
    this.state = { loading: undefined };
  }

  /**
   * Called after the component is mounted. You can do initialization that requires DOM nodes here. You can also make network requests here if you need to load data from a remote endpoint.
   */
  async componentDidMount() {
    this.setState({ ...(await this.getData()), loading: false });
  }

  /**
   * Override this method to fetch data for the widget.
   * @returns {Promise<Object>} The data for the widget.
   */
  async getData() {
    return {};
  }

  /**
   * Override this method to customize the styling of the widget.
   * @returns {string} The CSS class name for the widget.
   */
  styling() {
    return "";
  }

  /**
   * Override this method to define the header of the widget.
   * @returns {JSX.Element} The JSX element that defines the header of the widget.
   */
  header() {
    return null;
  }

  /**
   * Override this method to define the body of the widget.
   * @returns {JSX.Element} The JSX element that defines the body of the widget.
   */
  body() {
    return null;
  }

  /**
   * Override this method to define the footer of the widget.
   * @returns {JSX.Element} The JSX element that defines the footer of the widget.
   */
  footer() {
    return null;
  }

  /**
   * Renders the widget.
   * @returns {JSX.Element} The JSX element that defines the widget.
   */
  render() {
    return (
      <div className={mergeStyles(classNames.root, this.styling())}>
        {this.header() && <div className={classNames.header}>{this.header()}</div>}
        {this.body() && <div className={classNames.body}>{this.body()}</div>}
        {this.footer() && <div className={classNames.footer}>{this.footer()}</div>}
      </div>
    );
  }
}
