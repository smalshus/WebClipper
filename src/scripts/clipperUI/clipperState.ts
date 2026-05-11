import {ClientInfo} from "../clientInfo";
import {PageInfo} from "../pageInfo";
import {UserInfo} from "../userInfo";

import {SmartValue} from "../communicator/smartValue";

import {InvokeOptions} from "../extensions/invokeOptions";

import {ClipMode} from "./clipMode";
import {Status} from "./status";

// Minimal ClipperState type retained for clipperUrls.ts feedback URL
// generation. The full V1 ClipperState (with augmentation/full-page/PDF/
// bookmark result fields, ClipperInjectOptions, etc.) was removed alongside
// the V1 sidebar — the V3 renderer carries its own state in renderer.ts.
export interface DataResult<T> {
	data?: T;
	status: Status;
}

export interface ClipperState {
	uiExpanded?: boolean;
	fetchLocStringStatus?: Status;

	// Initialized at the start of the Clipper's instantiation
	invokeOptions?: InvokeOptions;

	// External "static" data
	userResult?: DataResult<UserInfo>;
	pageInfo?: PageInfo;
	clientInfo?: ClientInfo;

	// User input
	currentMode?: SmartValue<ClipMode>;
	saveLocation?: string;

	// Should be set when the Web Clipper enters a state that can not be recovered this session
	badState?: boolean;

	setState?: (partialState: ClipperState) => void;
	reset?: () => void;
}
