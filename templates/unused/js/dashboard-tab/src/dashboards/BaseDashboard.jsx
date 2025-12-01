import React, { Component } from "react";

import { mergeStyles } from "@fluentui/react";

/**
 * Returns the CSS class name for the dashboard.
 * @returns The CSS class name for the dashboard.
 * @internal
 */
function dashboardStyle(isMobile) {
  return mergeStyles({
    display: "grid",
    gap: "20px",
    padding: "20px",
    gridTemplateRows: "1fr",
    gridTemplateColumns: "4fr 6fr",
    ...(isMobile === true ? { gridTemplateColumns: "1fr", gridTemplateRows: "1fr" } : {}),
  });
}

/**
 * The base component that provides basic functionality to create a dashboard.
 */
export class BaseDashboard extends Component {
  /**
   * @internal
   */
  constructor(props) {
    super(props);
    this.state = {
      isMobile: undefined,
      showLogin: undefined,
      observer: undefined,
    };
    this.ref = React.createRef();
  }

  /**
   * Called after the component is mounted. You can do initialization that requires DOM nodes here. You can also make network requests here if you need to load data from a remote endpoint.
   */
  async componentDidMount() {
    // Observe the dashboard div for resize events
    const observer = new ResizeObserver((entries) => {
      for (const entry of entries) {
        if (entry.target === this.ref.current) {
          const { width } = entry.contentRect;
          this.setState({ isMobile: width < 600 });
        }
      }
    });
    observer.observe(this.ref.current);
    this.setState({ observer });
  }

  /**
   * Called before the component is unmounted and destroyed. You can do necessary cleanup here, such as invalidating timers, canceling network requests, or removing any DOM elements.
   */
  componentWillUnmount() {
    // Unobserve the dashboard div for resize events
    if (this.state.observer && this.ref.current) {
      this.state.observer.unobserve(this.ref.current);
    }
  }

  /**
   * Override this method to customize the styling of the dashboard.
   * @returns The CSS class name for the dashboard.
   */
  styling() {
    return "";
  }

  /**
   * Override this method to define the layout of the dashboard.
   * @returns The JSX element that defines the layout of the dashboard.
   */
  layout() {
    return null;
  }

  /**
   * Defines the default layout for the dashboard.
   */
  render() {
    return (
      <div
        ref={this.ref}
        className={mergeStyles(dashboardStyle(this.state.isMobile), this.styling())}
      >
        {this.layout()}
      </div>
    );
  }
}
