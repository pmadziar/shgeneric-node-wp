/// <reference path="../node_modules/@types/lodash/index.d.ts" />
/// <reference path="../node_modules/@types/sharepoint/index.d.ts" />
/// <reference path="../node_modules/@types/core-js/index.d.ts" />

import * as _ from "lodash"; 

export interface IDictionary<T> {
	[key: string]: T;
};

export class Helpers {
	public static guidRx: RegExp = new RegExp(`^[{(]{0,1}[a-fA-F0-9]{8}-{0,1}[a-fA-F0-9]{4}-{0,1}[a-fA-F0-9]{4}-{0,1}[a-fA-F0-9]{4}-{0,1}[a-fA-F0-9]{12}-{0,1}[})]{0,1}$`);
}


export function iterateSpCollection<T>(collection: SP.ClientObjectCollection<T>): Array<T> {
	let len = collection.get_count();
	let ret!: Array<T>;
	if (len > 0) {
		ret = new Array<T>();
		let collectionEnumerator = collection.getEnumerator();

		while (collectionEnumerator.moveNext()) {
			let value: T = collectionEnumerator.get_current();
			ret.push(value);
		}
	}
	return ret;
} 

export function newIgnoreErrorsPromise<T>(promise: Promise<T>) {
	return new Promise<T|null>((resolve: (value: T|null) => void, reject: (error: any) => void): void => {
		try {
			Promise.resolve(promise)
				.then(
					(value:T):void => { //resolve
						resolve(value);
					},(error: any): void => { //reject
						resolve(null);
					}
				);
		} catch (ex) {
			resolve(null);
		}
	});
}

export function cloneSpCamlQuery(query: SP.CamlQuery): SP.CamlQuery {
	let ret = new SP.CamlQuery;
	let datesInUtc = query.get_datesInUtc();
	if (typeof datesInUtc === 'boolean') {
		ret.set_datesInUtc(datesInUtc);
	}

	let folderServerRelativeUrl = query.get_folderServerRelativeUrl();
	if(!_.isEmpty(folderServerRelativeUrl)){
		ret.set_folderServerRelativeUrl(folderServerRelativeUrl);
	}

	let listItemCollectionPosition = query.get_listItemCollectionPosition();
	if (typeof listItemCollectionPosition !== 'undefined' && listItemCollectionPosition !== null){
		ret.set_listItemCollectionPosition(listItemCollectionPosition);
	}

	let viewXml = query.get_viewXml();
	if (!_.isEmpty(viewXml)) {
		ret.set_viewXml(viewXml);
	}

	return ret;
}
