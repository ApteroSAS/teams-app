declare const window;

export class Properties {
  getGlobalProperties<T>(): T {
    if (typeof window !== "undefined") {
      return window.APP_PROPS || {};
    } else {
      return {} as T;
    }
  }

  get<T>(key: string): T {
    return this.getGlobalProperties<any>()[key] as T;
  }
}
