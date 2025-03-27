export const delay = (delay: number) => {
  return new Promise((resolve) => setTimeout(resolve, delay));
};

type MaybePromise<T> = T | Promise<T>;
export function asPromise<T>(maybePromise: MaybePromise<T>): Promise<T> {
  if (typeof ((maybePromise as any)?.then) === 'function') {
    return maybePromise as Promise<T>;
  } else {
    return Promise.resolve(maybePromise);
  }
}
