/// <reference path="../node_modules/@types/lodash/index.d.ts" />
/// <reference path="../node_modules/@types/sharepoint/index.d.ts" />

export interface INvPromiseSvc<T> {
	GetAsync: () => Promise<INvPromiseSvc<T>>;
	ClientContext: SP.ClientContext;
	Site: INvPromiseSvc<SP.Site>;
	Web: INvPromiseSvc<SP.Web>;
	List: INvPromiseSvc<SP.List>;
	Target: T;
}
