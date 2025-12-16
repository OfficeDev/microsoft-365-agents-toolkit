import { useEffect, useReducer } from "react";

const createReducer = () => (state, action) => {
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
export function useData(fetchDataAsync, options) {
  const auto = options?.autoLoad ?? true;
  const [{ data, loading, error }, dispatch] = useReducer(createReducer(), {
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
