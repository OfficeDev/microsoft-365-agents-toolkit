import React, { Component } from "react";

import { mergeStyles } from "@fluentui/react";

/**
 * Returns the CSS class name for the dashboard.
 * @returns The CSS class name for the dashboard.
 * @internal
 */
function dashboardStyle(isMobile?: boolean) {
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
 * The state interface for the BaseDashboard component.
 */
interface BaseDashboardState {
  /**
   * A boolean property that indicates whether the dashboard layout should be optimized for mobile devices.
   */
  isMobile?: boolean;

  /**
   * A boolean property that indicates whether the login page should be displayed.
   */
  showLogin?: boolean;

  /**
   * The resize observer for the dashboard.
   * @internal
   */
  observer?: ResizeObserver;
}

/**
 * The base component that provides basic functionality to create a dashboard.
 * @typeParam P The type of props.
 * @typeParam S The type of state.
 */
export class BaseDashboard<P, S> extends Component<P, S & BaseDashboardState> {
  /**
   * @internal
   */
  private ref: React.RefObject<HTMLDivElement>;

  /**
   * Constructor of BaseDashboard.
   * @param {Readonly<P>} props The properties for the dashboard.
   */
  public constructor(props: Readonly<P>) {
    super(props);
    this.state = {
      isMobile: undefined,
      showLogin: undefined,
      observer: undefined,
    } as S & BaseDashboardState;
    this.ref = React.createRef<HTMLDivElement>();
  }

  /**
   * Called after the component is mounted. You can do initialization that requires DOM nodes here. You can also make network requests here if you need to load data from a remote endpoint.
   */
  public async componentDidMount() {
    // Observe the dashboard div for resize events
    const observer = new ResizeObserver((entries) => {
      for (const entry of entries) {
        if (entry.target === this.ref.current) {
          const { width } = entry.contentRect;
          this.setState({ isMobile: width < 600 } as S & BaseDashboardState);
        }
      }
    });
    observer.observe(this.ref.current!);
  }

  /**
   * Called before the component is unmounted and destroyed. You can do necessary cleanup here, such as invalidating timers, canceling network requests, or removing any DOM elements.
   */
  public componentWillUnmount(): void {
    // Unobserve the dashboard div for resize events
    if (this.state.observer && this.ref.current) {
      this.state.observer.unobserve(this.ref.current);
    }
  }

  /**
   * Defines the default layout for the dashboard.
   */
  public render() {
    return (
      <div
        ref={this.ref}
        className={mergeStyles(dashboardStyle(this.state.isMobile), this.styling())}
      >
        {this.layout()}
      </div>
    );
  }

  /**
   * Override this method to define the layout of the widget in the dashboard.
   * @returns The layout of the widget in the dashboard.
   * @public
   */
  protected layout(): JSX.Element | undefined {
    return undefined;
  }

  /**
   * Override this method to customize the dashboard style.
   * @returns The className for customizing the dashboard style.
   * @public
   */
  protected styling(): string {
    return null;
  }
}
