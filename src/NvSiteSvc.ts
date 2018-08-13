/// <reference path="../node_modules/@types/lodash/index.d.ts" />
/// <reference path="../node_modules/@types/sharepoint/index.d.ts" />
/// <reference path="../node_modules/@types/core-js/index.d.ts" />


import * as _ from "lodash"; 
import { INvPromiseSvc } from "./INvPromiseSvc";

	export class NvSiteSvc implements INvPromiseSvc<SP.Site> {
		private siteServerRelativeUrl?: string;;
		private _site!: SP.Site;

        constructor(siteServerRelativeUrl?: string) {
            if (!_.isEmpty(siteServerRelativeUrl)) {
                this.siteServerRelativeUrl = siteServerRelativeUrl;
            }
        }

		GetAsync: () => Promise<INvPromiseSvc<SP.Site>> = (): Promise<INvPromiseSvc<SP.Site>> => {
			return new Promise<INvPromiseSvc<SP.Site>>((resolve: (rsite: Promise<NvSiteSvc>) => void, reject: (error: any) => void): void => {

	 			if (_.isEmpty(this.siteServerRelativeUrl)) {
	 				this.ClientContext = SP.ClientContext.get_current();
	 			} else {
	 				this.ClientContext = new SP.ClientContext(this.siteServerRelativeUrl);
	 			}
	 			this._site = this.ClientContext.get_site();
				this.ClientContext.load(this._site);

	 			this.ClientContext.executeQueryAsync(
	 				() :void => {
                         this.Target = this._site;
					     this.Site = this;
						resolve(Promise.resolve(this));
	 				},
					(sender: any, args: SP.ClientRequestFailedEventArgs): void => {
						let error = new Error(args.get_message());
						reject(error);
	 				}
	 			);
			});
		};

		public ClientContext!: SP.ClientContext;
		public Site!: INvPromiseSvc<SP.Site>;
		public Web!: INvPromiseSvc<SP.Web>;
		public List!: INvPromiseSvc<SP.List>;
		public Target!: SP.Site;
		}
