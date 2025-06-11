import { useEffect, useReducer } from "react";

type State<T> = {
  /**
   * User data.
   */
  data?: T;
  /**
   * Status of data loading.
   */
  loading: boolean;
  /**
   * Error information.
   */
  error?: unknown;
};

type Action<T> =
  | { type: "loading" }
  | { type: "result"; result: T }
  | { type: "error"; error: unknown };

export type Data<T> = State<T> & {
  /**
   * reload function.
   */
  reload: () => void;
};

const createReducer =
  <T>() =>
  (state: State<T>, action: Action<T>): State<T> => {
    switch (action.type) {
      case "loading":
        return { data: state.data, loading: true };
      case "result":
        return { data: action.result, loading: false };
      case "error":
        return { loading: false, error: action.error };
    }
  };

/**
 * Helper function to fetch data with status and error.
 *
 * @param fetchDataAsync - async function of how to fetch data
 * @param options - if autoLoad is true, reload data immediately
 * @returns data, loading status, error and reload function
 *
 * @public
 */
export function useData<T>(
  fetchDataAsync: () => Promise<T>,
  options?: { autoLoad: boolean }
): Data<T> {
  const auto = options?.autoLoad ?? true;
  const [{ data, loading, error }, dispatch] = useReducer(createReducer<T>(), {
    loading: auto,
  });
  function reload() {
    if (!loading) dispatch({ type: "loading" });
    fetchDataAsync()
      .then((data) => dispatch({ type: "result", result: data }))
      .catch((error) => dispatch({ type: "error", error }));
  }
  useEffect(() => {
    if (auto) reload();
  }, []);
  return { data, loading, error, reload };
}
