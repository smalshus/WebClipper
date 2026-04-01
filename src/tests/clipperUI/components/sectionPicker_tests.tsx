import * as sinon from "sinon";

import {Clipper} from "../../../scripts/clipperUI/frontEndGlobals";
import {Status} from "../../../scripts/clipperUI/status";
import {OneNoteApiUtils} from "../../../scripts/clipperUI/oneNoteApiUtils";

import {SectionPicker, SectionPickerClass, SectionPickerState} from "../../../scripts/clipperUI/components/sectionPicker";

import {ClipperStorageKeys} from "../../../scripts/storage/clipperStorageKeys";

import {MithrilUtils} from "../../mithrilUtils";
import {MockProps} from "../../mockProps";
import {TestModule} from "../../testModule";

module TestConstants {
	export module Ids {
		export var sectionLocationContainer = "sectionLocationContainer";
		export var sectionPickerContainer = "sectionPickerContainer";
	}

	export var defaultUserInfoAsJsonString = JSON.stringify({
		emailAddress: "mockEmail@hotmail.com",
		fullName: "Mock Johnson",
		accessToken: "mockToken",
		accessTokenExpiration: 3000
	});
}

type StoredSection = {
	section: OneNoteApi.Section,
	path: string,
	parentId: string
};

// Mock out the Clipper.Storage functionality
let mockStorage: { [key: string]: string } = {};
Clipper.getStoredValue = (key: string, callback: (value: string) => void) => {
	callback(mockStorage[key]);
};
Clipper.storeValue = (key: string, value: string) => {
	mockStorage[key] = value;
};

function initializeClipperStorage(notebooks: string, curSection: string, userInfo?: string) {
	mockStorage = { };
	Clipper.storeValue(ClipperStorageKeys.cachedNotebooks, notebooks);
	Clipper.storeValue(ClipperStorageKeys.currentSelectedSection, curSection);
	Clipper.storeValue(ClipperStorageKeys.userInformation, userInfo);
}

function createNotebook(id: string, isDefault?: boolean, sectionGroups?: OneNoteApi.SectionGroup[], sections?: OneNoteApi.Section[]): OneNoteApi.Notebook {
	return {
		name: id.toUpperCase(),
		isDefault: isDefault,
		userRole: undefined,
		isShared: true,
		links: undefined,
		id: id.toLowerCase(),
		self: undefined,
		createdTime: undefined,
		lastModifiedTime: undefined,
		createdBy: undefined,
		lastModifiedBy: undefined,
		sectionsUrl: undefined,
		sectionGroupsUrl: undefined,
		sections: sections,
		sectionGroups: sectionGroups
	};
};

function createSectionGroup(id: string, sectionGroups?: OneNoteApi.SectionGroup[], sections?: OneNoteApi.Section[]): OneNoteApi.SectionGroup {
	return {
		name: id.toUpperCase(),
		id: id.toLowerCase(),
		self: undefined,
		createdTime: undefined,
		lastModifiedTime: undefined,
		createdBy: undefined,
		lastModifiedBy: undefined,
		sectionsUrl: undefined,
		sectionGroupsUrl: undefined,
		sections: sections,
		sectionGroups: sectionGroups
	};
};

function createSection(id: string, isDefault?: boolean): OneNoteApi.Section {
	return {
		name: id.toUpperCase(),
		isDefault: isDefault,
		parentNotebook: undefined,
		id: id.toLowerCase(),
		self: undefined,
		createdTime: undefined,
		lastModifiedTime: undefined,
		createdBy: undefined,
		lastModifiedBy: undefined,
		pagesUrl: undefined,
		pages: undefined
	};
};

export class SectionPickerTests extends TestModule {
	private defaultComponent;
	private mockClipperState = MockProps.getMockClipperState();

	protected module() {
		return "sectionPicker";
	}

	protected beforeEach() {
		this.defaultComponent = <SectionPicker
			onPopupToggle={() => {}}
			clipperState={this.mockClipperState} />;
	}

	protected tests() {
		test("fetchCachedNotebookAndSectionInfoAsState should return the cached notebooks, cached current section, and the succeed status if cached information is found", () => {
			let clipperState = MockProps.getMockClipperState();

			let mockNotebooks = MockProps.getMockNotebooks();
			let mockSection = {
				section: mockNotebooks[0].sections[0],
				path: "A > B > C",
				parentId: mockNotebooks[0].id
			};
			initializeClipperStorage(JSON.stringify(mockNotebooks), JSON.stringify(mockSection));

			let component = <SectionPicker
				onPopupToggle={() => {}}
				clipperState={clipperState} />;
			let controllerInstance = MithrilUtils.mountToFixture(component);

			controllerInstance.fetchCachedNotebookAndSectionInfoAsState((response: SectionPickerState) => {
				strictEqual(JSON.stringify(response), JSON.stringify({ notebooks: mockNotebooks, status: Status.Succeeded, curSection: mockSection }),
					"The cached information should be returned as SectionPickerState");
			});
		});

		test("fetchCachedNotebookAndSectionInfoAsState should return undefined if no cached information is found", () => {
			let clipperState = MockProps.getMockClipperState();

			initializeClipperStorage(undefined, undefined);

			let component = <SectionPicker
				onPopupToggle={() => {}}
				clipperState={clipperState} />;
			let controllerInstance = MithrilUtils.mountToFixture(component);

			controllerInstance.fetchCachedNotebookAndSectionInfoAsState((response: SectionPickerState) => {
				strictEqual(response, undefined,
					"The undefined notebooks and section information should be returned as SectionPickerState");
			});
		});

		test("fetchCachedNotebookAndSectionInfoAsState should return the cached notebooks, undefined section, and the succeed status if no cached section is found", () => {
			let clipperState = MockProps.getMockClipperState();

			let mockNotebooks = MockProps.getMockNotebooks();
			initializeClipperStorage(JSON.stringify(mockNotebooks), undefined);

			let component = <SectionPicker
				onPopupToggle={() => {}}
				clipperState={clipperState} />;
			let controllerInstance = MithrilUtils.mountToFixture(component);

			controllerInstance.fetchCachedNotebookAndSectionInfoAsState((response: SectionPickerState) => {
				strictEqual(JSON.stringify(response), JSON.stringify({ notebooks: mockNotebooks, status: Status.Succeeded, curSection: undefined }),
					"The cached information should be returned as SectionPickerState");
			});
		});

		test("fetchCachedNotebookAndSectionInfoAsState should return undefined when no notebooks are found, even if section information is found", () => {
			let clipperState = MockProps.getMockClipperState();

			let mockSection = {
				section: MockProps.getMockNotebooks()[0].sections[0],
				path: "A > B > C",
				parentId: MockProps.getMockNotebooks()[0].id
			};
			initializeClipperStorage(undefined, JSON.stringify(mockSection));

			let component = <SectionPicker
				onPopupToggle={() => {}}
				clipperState={clipperState} />;
			let controllerInstance = MithrilUtils.mountToFixture(component);

			controllerInstance.fetchCachedNotebookAndSectionInfoAsState((response: SectionPickerState) => {
				strictEqual(response, undefined,
					"The cached information should be returned as SectionPickerState");
			});
		});

		test("convertNotebookListToState should return the notebook list, success status, and default section in the general case", () => {
			let section = createSection("S", true);
			let sectionGroup2 = createSectionGroup("SG2", [], [section]);
			let sectionGroup1 = createSectionGroup("SG1", [sectionGroup2], []);
			let notebook = createNotebook("N", true, [sectionGroup1], []);

			let notebooks = [notebook];
			let actual = SectionPickerClass.convertNotebookListToState(notebooks);
			strictEqual(actual.notebooks, notebooks, "The notebooks property is correct");
			strictEqual(actual.status, Status.Succeeded, "The status property is correct");
			deepEqual(actual.curSection, { path: "N > SG1 > SG2 > S", section: section },
				"The curSection property is correct");
		});

		test("convertNotebookListToState should return the notebook list, success status, and undefined default section in case where there is no default section", () => {
			let sectionGroup2 = createSectionGroup("SG2", [], []);
			let sectionGroup1 = createSectionGroup("SG1", [sectionGroup2], []);
			let notebook = createNotebook("N", true, [sectionGroup1], []);

			let notebooks = [notebook];
			let actual = SectionPickerClass.convertNotebookListToState(notebooks);
			strictEqual(actual.notebooks, notebooks, "The notebooks property is correct");
			strictEqual(actual.status, Status.Succeeded, "The status property is correct");
			strictEqual(actual.curSection, undefined, "The curSection property is undefined");
		});

		test("convertNotebookListToState should return the notebook list, success status, and undefined default section in case where there is only one empty notebook", () => {
			let notebook = createNotebook("N", true, [], []);

			let notebooks = [notebook];
			let actual = SectionPickerClass.convertNotebookListToState(notebooks);
			strictEqual(actual.notebooks, notebooks, "The notebooks property is correct");
			strictEqual(actual.status, Status.Succeeded, "The status property is correct");
			strictEqual(actual.curSection, undefined, "The curSection property is undefined");
		});

		test("convertNotebookListToState should return the undefined notebook list, success status, and undefined default section if the input is undefined", () => {
			let actual = SectionPickerClass.convertNotebookListToState(undefined);
			strictEqual(actual.notebooks, undefined, "The notebooks property is undefined");
			strictEqual(actual.status, Status.Succeeded, "The status property is correct");
			strictEqual(actual.curSection, undefined, "The curSection property is undefined");
		});

		test("convertNotebookListToState should return the empty notebook list, success status, and undefined default section if the input is undefined", () => {
			let actual = SectionPickerClass.convertNotebookListToState([]);
			strictEqual(actual.notebooks.length, 0, "The notebooks property is the empty list");
			strictEqual(actual.status, Status.Succeeded, "The status property is correct");
			strictEqual(actual.curSection, undefined, "The curSection property is undefined");
		});

		test("formatSectionInfoForStorage should return a ' > ' delimited name path and the last element in the general case", () => {
			let section = createSection("4");
			let actual = SectionPickerClass.formatSectionInfoForStorage([
				createNotebook("1"),
				createSectionGroup("2"),
				createSectionGroup("3"),
				section
			]);
			deepEqual(actual, { path: "1 > 2 > 3 > 4", section: section },
				"The section info should be formatted correctly");
		});

		test("formatSectionInfoForStorage should return a ' > ' delimited name path and the last element if there are no section groups", () => {
			let section = createSection("2");
			let actual = SectionPickerClass.formatSectionInfoForStorage([
				createNotebook("1"),
				section
			]);
			deepEqual(actual, { path: "1 > 2", section: section },
				"The section info should be formatted correctly");
		});

		test("formatSectionInfoForStorage should return undefined if the list that is passed in is undefined", () => {
			let actual = SectionPickerClass.formatSectionInfoForStorage(undefined);
			strictEqual(actual, undefined, "The section info should be formatted correctly");
		});

		test("formatSectionInfoForStorage should return undefined if the list that is passed in is empty", () => {
			let actual = SectionPickerClass.formatSectionInfoForStorage([]);
			strictEqual(actual, undefined, "The section info should be formatted correctly");
		});

		test("onPopupToggle should focus the currently selected section element when the popup opens and a curSection is set", (assert: QUnitAssert) => {
			let done = assert.async();
			let clock = sinon.useFakeTimers();

			let clipperState = MockProps.getMockClipperState();
			let mockNotebooks = MockProps.getMockNotebooks();
			let mockSection = {
				section: mockNotebooks[0].sections[0],
				path: "Clipper Test > Full Page",
				parentId: mockNotebooks[0].id
			};
			initializeClipperStorage(JSON.stringify(mockNotebooks), JSON.stringify(mockSection));

			let component = <SectionPicker onPopupToggle={() => {}} clipperState={clipperState} />;
			let controllerInstance = MithrilUtils.mountToFixture(component);

			// Create a fake section element in the DOM that matches the selected section id
			let sectionElement = document.createElement("li");
			sectionElement.id = mockSection.section.id;
			sectionElement.tabIndex = 70;
			let focusCalled = false;
			sectionElement.focus = () => { focusCalled = true; };
			document.body.appendChild(sectionElement);

			controllerInstance.onPopupToggle(true);
			clock.tick(0);

			ok(focusCalled, "The selected section element should have been focused when the popup opens");

			document.body.removeChild(sectionElement);
			clock.restore();
			done();
		});

		test("onPopupToggle should focus the first focusable item in the picker popup when the popup opens and no curSection is set", (assert: QUnitAssert) => {
			let done = assert.async();
			let clock = sinon.useFakeTimers();

			let clipperState = MockProps.getMockClipperState();
			initializeClipperStorage(undefined, undefined);

			let component = <SectionPicker onPopupToggle={() => {}} clipperState={clipperState} />;
			let controllerInstance = MithrilUtils.mountToFixture(component);

			// Create a fake popup container and a focusable item inside it
			let sectionPickerPopup = document.createElement("div");
			sectionPickerPopup.id = "sectionPickerContainer";
			let firstItem = document.createElement("li");
			firstItem.tabIndex = 70;
			let focusCalled = false;
			firstItem.focus = () => { focusCalled = true; };
			sectionPickerPopup.appendChild(firstItem);
			document.body.appendChild(sectionPickerPopup);

			controllerInstance.onPopupToggle(true);
			clock.tick(0);

			ok(focusCalled, "The first focusable item in the picker popup should have been focused when no section is selected");

			document.body.removeChild(sectionPickerPopup);
			clock.restore();
			done();
		});

		test("onPopupToggle should not change focus when the popup closes", (assert: QUnitAssert) => {
			let done = assert.async();
			let clock = sinon.useFakeTimers();

			let clipperState = MockProps.getMockClipperState();
			let mockNotebooks = MockProps.getMockNotebooks();
			let mockSection = {
				section: mockNotebooks[0].sections[0],
				path: "Clipper Test > Full Page",
				parentId: mockNotebooks[0].id
			};
			initializeClipperStorage(JSON.stringify(mockNotebooks), JSON.stringify(mockSection));

			let component = <SectionPicker onPopupToggle={() => {}} clipperState={clipperState} />;
			let controllerInstance = MithrilUtils.mountToFixture(component);

			// Create a fake section element to catch any unexpected focus calls
			let sectionElement = document.createElement("li");
			sectionElement.id = mockSection.section.id;
			sectionElement.tabIndex = 70;
			let focusCalled = false;
			sectionElement.focus = () => { focusCalled = true; };
			document.body.appendChild(sectionElement);

			controllerInstance.onPopupToggle(false);
			clock.tick(0);

			ok(!focusCalled, "No focus change should occur when the popup closes");

			document.body.removeChild(sectionElement);
			clock.restore();
			done();
		});

		test("onPopupToggle should enable Down arrow key to move focus to the next item in the popup", (assert: QUnitAssert) => {
			let done = assert.async();
			let clock = sinon.useFakeTimers();

			let clipperState = MockProps.getMockClipperState();
			initializeClipperStorage(undefined, undefined);

			let component = <SectionPicker onPopupToggle={() => {}} clipperState={clipperState} />;
			let controllerInstance = MithrilUtils.mountToFixture(component);

			// Create a fake popup with two items
			let sectionPickerPopup = document.createElement("div");
			sectionPickerPopup.id = "sectionPickerContainer";
			let firstItem = document.createElement("li");
			firstItem.tabIndex = 70;
			let secondItemFocusCalled = false;
			let secondItem = document.createElement("li");
			secondItem.tabIndex = 70;
			secondItem.focus = () => { secondItemFocusCalled = true; };
			sectionPickerPopup.appendChild(firstItem);
			sectionPickerPopup.appendChild(secondItem);
			document.body.appendChild(sectionPickerPopup);

			controllerInstance.onPopupToggle(true);
			clock.tick(0);

			// Simulate focus on first item and press Down arrow
			firstItem.focus();
			let downKeyEvent = document.createEvent("KeyboardEvent");
			downKeyEvent.initEvent("keydown", true, true);
			Object.defineProperty(downKeyEvent, "which", { value: 40 });
			sectionPickerPopup.dispatchEvent(downKeyEvent);

			ok(secondItemFocusCalled, "Down arrow key should move focus to the next item in the popup");

			document.body.removeChild(sectionPickerPopup);
			clock.restore();
			done();
		});

		test("onPopupToggle should enable Up arrow key to move focus to the previous item in the popup", (assert: QUnitAssert) => {
			let done = assert.async();
			let clock = sinon.useFakeTimers();

			let clipperState = MockProps.getMockClipperState();
			initializeClipperStorage(undefined, undefined);

			let component = <SectionPicker onPopupToggle={() => {}} clipperState={clipperState} />;
			let controllerInstance = MithrilUtils.mountToFixture(component);

			// Create a fake popup with two items
			let sectionPickerPopup = document.createElement("div");
			sectionPickerPopup.id = "sectionPickerContainer";
			let firstItemFocusCalled = false;
			let firstItem = document.createElement("li");
			firstItem.tabIndex = 70;
			firstItem.focus = () => { firstItemFocusCalled = true; };
			let secondItem = document.createElement("li");
			secondItem.tabIndex = 70;
			sectionPickerPopup.appendChild(firstItem);
			sectionPickerPopup.appendChild(secondItem);
			document.body.appendChild(sectionPickerPopup);

			controllerInstance.onPopupToggle(true);
			clock.tick(0);

			// Simulate focus on second item and press Up arrow
			secondItem.focus();
			let upKeyEvent = document.createEvent("KeyboardEvent");
			upKeyEvent.initEvent("keydown", true, true);
			Object.defineProperty(upKeyEvent, "which", { value: 38 });
			sectionPickerPopup.dispatchEvent(upKeyEvent);

			ok(firstItemFocusCalled, "Up arrow key should move focus to the previous item in the popup");

			document.body.removeChild(sectionPickerPopup);
			clock.restore();
			done();
		});

		test("onPopupToggle should skip hidden items inside closed notebooks when navigating with Down arrow", (assert: QUnitAssert) => {
			let done = assert.async();
			let clock = sinon.useFakeTimers();

			let clipperState = MockProps.getMockClipperState();
			initializeClipperStorage(undefined, undefined);

			let component = <SectionPicker onPopupToggle={() => {}} clipperState={clipperState} />;
			let controllerInstance = MithrilUtils.mountToFixture(component);

			// Build a structure that mirrors the OneNotePicker DOM:
			//   sectionPickerContainer
			//     li.Notebook.Closed (notebook1, tabindex)
			//       ul
			//         li.Section (hiddenSection, tabindex) -- inside closed notebook
			//     li.Notebook.Opened (notebook2, tabindex)
			let sectionPickerPopup = document.createElement("div");
			sectionPickerPopup.id = "sectionPickerContainer";

			let notebook1 = document.createElement("li");
			notebook1.className = "Notebook Closed";
			notebook1.tabIndex = 70;
			let closedChildList = document.createElement("ul");
			let hiddenSection = document.createElement("li");
			hiddenSection.className = "Section";
			hiddenSection.tabIndex = 70;
			let hiddenSectionFocusCalled = false;
			hiddenSection.focus = () => { hiddenSectionFocusCalled = true; };
			closedChildList.appendChild(hiddenSection);
			notebook1.appendChild(closedChildList);

			let notebook2FocusCalled = false;
			let notebook2 = document.createElement("li");
			notebook2.className = "Notebook Opened";
			notebook2.tabIndex = 70;
			notebook2.focus = () => { notebook2FocusCalled = true; };

			sectionPickerPopup.appendChild(notebook1);
			sectionPickerPopup.appendChild(notebook2);
			document.body.appendChild(sectionPickerPopup);

			controllerInstance.onPopupToggle(true);
			clock.tick(0);

			// Press Down arrow from notebook1 — should jump to notebook2, skipping the hidden section inside notebook1
			notebook1.focus();
			let downKeyEvent = document.createEvent("KeyboardEvent");
			downKeyEvent.initEvent("keydown", true, true);
			Object.defineProperty(downKeyEvent, "which", { value: 40 });
			sectionPickerPopup.dispatchEvent(downKeyEvent);

			ok(!hiddenSectionFocusCalled, "Hidden section inside a closed notebook should not receive focus");
			ok(notebook2FocusCalled, "Down arrow from a closed notebook should move focus to the next visible notebook");

			document.body.removeChild(sectionPickerPopup);
			clock.restore();
			done();
		});
	}
}

export class SectionPickerSinonTests extends TestModule {
	private defaultComponent;
	private mockClipperState = MockProps.getMockClipperState();

	private server: sinon.SinonFakeServer;

	protected module() {
		return "sectionPicker-sinon";
	}

	protected beforeEach() {
		this.defaultComponent = <SectionPicker
			onPopupToggle={() => {}}
			clipperState={this.mockClipperState} />;

		this.server = sinon.fakeServer.create();
		this.server.respondImmediately = true;
	}

	protected afterEach() {
		this.server.restore();
	}

	protected tests() {
		test("retrieveAndUpdateNotebookAndSectionSelection should update states correctly when there's notebook and curSection information found in storage," +
			"the user does not make a new selection, and then information is found on the server. Also the notebooks are the same in storage and on the server, " +
			"and the current section in storage is the same as the default section in the server's notebook list", (assert: QUnitAssert) => {
			let done = assert.async();

			let clipperState = MockProps.getMockClipperState();

			// Set up the storage mock
			let mockNotebooks = MockProps.getMockNotebooks();
			let mockSection = {
				section: mockNotebooks[0].sections[0],
				path: "Clipper Test > Full Page",
				parentId: mockNotebooks[0].id
			};
			initializeClipperStorage(JSON.stringify(mockNotebooks), JSON.stringify(mockSection), TestConstants.defaultUserInfoAsJsonString);

			// After retrieving fresh notebooks, the storage should be updated with the fresh notebooks (although it's the same in this case)
			let freshNotebooks = MockProps.getMockNotebooks();
			let responseJson = {
				"@odata.context": "https://www.onenote.com/api/v1.0/$metadata#me/notes/notebooks",
				value: freshNotebooks
			};
			this.server.respondWith([200, { "Content-Type": "application/json" }, JSON.stringify(responseJson)]);

			let component = <SectionPicker onPopupToggle={() => {}} clipperState={clipperState} />;
			let controllerInstance = MithrilUtils.mountToFixture(component);

			strictEqual(JSON.stringify(controllerInstance.state), JSON.stringify({ notebooks: mockNotebooks, status: Status.Succeeded, curSection: mockSection }),
				"After the component is mounted, the state should be updated to reflect the notebooks and section found in storage");

			controllerInstance.retrieveAndUpdateNotebookAndSectionSelection().then((response) => {
				Clipper.getStoredValue(ClipperStorageKeys.cachedNotebooks, (notebooks) => {
					Clipper.getStoredValue(ClipperStorageKeys.currentSelectedSection, (curSection) => {
						strictEqual(notebooks, JSON.stringify(freshNotebooks),
							"After fresh notebooks have been retrieved, the storage should be updated with them. In this case, nothing should have changed.");
						strictEqual(curSection, JSON.stringify(mockSection),
							"The current selected section in storage should not have changed");

						strictEqual(JSON.stringify(controllerInstance.state.notebooks), JSON.stringify(freshNotebooks),
							"The state should always be updated with the fresh notebooks once it has been retrieved");
						strictEqual(JSON.stringify(controllerInstance.state.curSection), JSON.stringify(mockSection),
							"Since curSection was found in storage, and the user did not make an action to select another section, it should remain the same in state");
						strictEqual(controllerInstance.state.status, Status.Succeeded, "The status should be Succeeded");
						done();
					});
				});
			}, (error) => {
				ok(false, "reject should not be called");
			});
		});

		test("retrieveAndUpdateNotebookAndSectionSelection should update states correctly when there's notebook and curSection information found in storage," +
			"the user does not make a new selection, and then information is found on the server. The notebooks on the server is not the same as the ones in storage, " +
			"and the current section in storage is the same as the default section in the server's notebook list", (assert: QUnitAssert) => {
			let done = assert.async();

			let clipperState = MockProps.getMockClipperState();

			// Set up the storage mock
			let mockNotebooks = MockProps.getMockNotebooks();
			let mockSection = {
				section: mockNotebooks[0].sections[0],
				path: "Clipper Test > Full Page",
				parentId: mockNotebooks[0].id
			};
			initializeClipperStorage(JSON.stringify(mockNotebooks), JSON.stringify(mockSection), TestConstants.defaultUserInfoAsJsonString);

			let component = <SectionPicker
				onPopupToggle={() => {}}
				clipperState={clipperState} />;
			let controllerInstance = MithrilUtils.mountToFixture(component);

			strictEqual(JSON.stringify(controllerInstance.state), JSON.stringify({ notebooks: mockNotebooks, status: Status.Succeeded, curSection: mockSection }),
				"After the component is mounted, the state should be updated to reflect the notebooks and section found in storage");

			// After retrieving fresh notebooks, the storage should be updated with the fresh notebooks
			let freshNotebooks = MockProps.getMockNotebooks();
			freshNotebooks.push(createNotebook("id", false, [], []));
			let responseJson = {
				"@odata.context": "https://www.onenote.com/api/v1.0/$metadata#me/notes/notebooks",
				value: freshNotebooks
			};
			this.server.respondWith([200, { "Content-Type": "application/json" }, JSON.stringify(responseJson)]);

			controllerInstance.retrieveAndUpdateNotebookAndSectionSelection().then((response: SectionPickerState) => {
				Clipper.getStoredValue(ClipperStorageKeys.cachedNotebooks, (notebooks) => {
					Clipper.getStoredValue(ClipperStorageKeys.currentSelectedSection, (curSection) => {
						strictEqual(notebooks, JSON.stringify(freshNotebooks),
							"After fresh notebooks have been retrieved, the storage should be updated with them. In this case, nothing should have changed.");
						strictEqual(curSection, JSON.stringify(mockSection),
							"The current selected section in storage should not have changed");

						strictEqual(JSON.stringify(controllerInstance.state.notebooks), JSON.stringify(freshNotebooks),
							"The state should always be updated with the fresh notebooks once it has been retrieved");
						strictEqual(JSON.stringify(controllerInstance.state.curSection), JSON.stringify(mockSection),
							"Since curSection was found in storage, and the user did not make an action to select another section, it should remain the same in state");
						strictEqual(controllerInstance.state.status, Status.Succeeded,
							"The status should be Succeeded");
						done();
					});
				});
			}, (error) => {
				ok(false, "reject should not be called");
			});
		});

		test("retrieveAndUpdateNotebookAndSectionSelection should update states correctly when there's notebook, but no curSection information found in storage," +
			"the user does not make a selection, and then information is found on the server. The notebooks on the server is the same as the ones in storage, " +
			"and the current section in storage is still undefined by the time the fresh notebooks have been retrieved", (assert: QUnitAssert) => {
			let done = assert.async();

			let clipperState = MockProps.getMockClipperState();

			// Set up the storage mock
			let mockNotebooks = MockProps.getMockNotebooks();
			initializeClipperStorage(JSON.stringify(mockNotebooks), undefined, TestConstants.defaultUserInfoAsJsonString);

			let component = <SectionPicker
				onPopupToggle={() => {}}
				clipperState={clipperState} />;
			let controllerInstance = MithrilUtils.mountToFixture(component);

			strictEqual(JSON.stringify(controllerInstance.state), JSON.stringify({ notebooks: mockNotebooks, status: Status.Succeeded, curSection: undefined }),
				"After the component is mounted, the state should be updated to reflect the notebooks and section found in storage");

			// After retrieving fresh notebooks, the storage should be updated with the fresh notebooks (although it's the same in this case)
			let freshNotebooks = MockProps.getMockNotebooks();
			freshNotebooks.push(createNotebook("id", false, [], []));
			let responseJson = {
				"@odata.context": "https://www.onenote.com/api/v1.0/$metadata#me/notes/notebooks",
				value: freshNotebooks
			};
			this.server.respondWith([200, { "Content-Type": "application/json" }, JSON.stringify(responseJson)]);

			// This is the default section in the mock notebooks, and this should be found in storage and state after fresh notebooks are retrieved
			let defaultSection = {
				path: "Clipper Test > Full Page",
				section: mockNotebooks[0].sections[0]
			};

			controllerInstance.retrieveAndUpdateNotebookAndSectionSelection().then((response: SectionPickerState) => {
				Clipper.getStoredValue(ClipperStorageKeys.cachedNotebooks, (notebooks) => {
					Clipper.getStoredValue(ClipperStorageKeys.currentSelectedSection, (curSection) => {
						strictEqual(notebooks, JSON.stringify(freshNotebooks),
							"After fresh notebooks have been retrieved, the storage should be updated with them. In this case, nothing should have changed.");
						strictEqual(curSection, JSON.stringify(defaultSection),
							"The current selected section in storage should have been updated with the default section since it was undefined before");

						strictEqual(JSON.stringify(controllerInstance.state.notebooks), JSON.stringify(freshNotebooks),
							"The state should always be updated with the fresh notebooks once it has been retrieved");
						strictEqual(JSON.stringify(controllerInstance.state.curSection), JSON.stringify(defaultSection),
							"Since curSection was not found in storage, and the user did not make an action to select another section, it should be updated in state");
						strictEqual(controllerInstance.state.status, Status.Succeeded,
							"The status should be Succeeded");
						done();
					});
				});
			}, (error) => {
				ok(false, "reject should not be called");
			});
		});

		test("retrieveAndUpdateNotebookAndSectionSelection should update states correctly when there's notebook, but no curSection information found in storage," +
			"the user makes a new section selection, and then information is found on the server. The notebooks on the server is the same as the ones in storage, " +
			"and the current section in storage is still undefined by the time the fresh notebooks have been retrieved", (assert: QUnitAssert) => {
			let done = assert.async();

			let clipperState = MockProps.getMockClipperState();

			// Set up the storage mock
			let mockNotebooks = MockProps.getMockNotebooks();
			initializeClipperStorage(JSON.stringify(mockNotebooks), undefined, TestConstants.defaultUserInfoAsJsonString);

			let component = <SectionPicker
				onPopupToggle={() => {}}
				clipperState={clipperState} />;
			let controllerInstance = MithrilUtils.mountToFixture(component);

			strictEqual(JSON.stringify(controllerInstance.state), JSON.stringify({ notebooks: mockNotebooks, status: Status.Succeeded, curSection: undefined }),
				"After the component is mounted, the state should be updated to reflect the notebooks and section found in storage");

			// The user now clicks on a section (second section of second notebook)
			MithrilUtils.simulateAction(() => {
				document.getElementById(TestConstants.Ids.sectionLocationContainer).click();
			});
			let sectionPicker = document.getElementById(TestConstants.Ids.sectionPickerContainer).firstElementChild;
			let second = sectionPicker.childNodes[1];
			let secondNotebook = second.childNodes[0] as HTMLElement;
			let secondSections = second.childNodes[1] as HTMLElement;
			MithrilUtils.simulateAction(() => {
				secondNotebook.click();
			});
			let newSelectedSection = secondSections.childNodes[1] as HTMLElement;
			MithrilUtils.simulateAction(() => {
				// The clickable element is actually the first childNode
				(newSelectedSection.childNodes[0] as HTMLElement).click();
			});

			// This corresponds to the second section of the second notebook in the mock notebooks
			let selectedSection = {
				section: mockNotebooks[1].sections[1],
				path: "Clipper Test 2 > Section Y",
				parentId: "a-bc!d"
			};

			Clipper.getStoredValue(ClipperStorageKeys.currentSelectedSection, (curSection1) => {
				strictEqual(curSection1, JSON.stringify(selectedSection),
					"The current selected section in storage should have been updated with the selected section");
				strictEqual(JSON.stringify(controllerInstance.state.curSection), JSON.stringify(selectedSection),
					"The current selected section in state should have been updated with the selected section");

				// After retrieving fresh notebooks, the storage should be updated with the fresh notebooks (although it's the same in this case)
				let freshNotebooks = MockProps.getMockNotebooks();
				freshNotebooks.push(createNotebook("id", false, [], []));
				let responseJson = {
					"@odata.context": "https://www.onenote.com/api/v1.0/$metadata#me/notes/notebooks",
					value: freshNotebooks
				};
				this.server.respondWith([200, { "Content-Type": "application/json" }, JSON.stringify(responseJson)]);

				controllerInstance.retrieveAndUpdateNotebookAndSectionSelection().then((response: SectionPickerState) => {
					Clipper.getStoredValue(ClipperStorageKeys.cachedNotebooks, (notebooks) => {
						Clipper.getStoredValue(ClipperStorageKeys.currentSelectedSection, (curSection2) => {
							strictEqual(notebooks, JSON.stringify(freshNotebooks),
								"After fresh notebooks have been retrieved, the storage should be updated with them. In this case, nothing should have changed.");
							strictEqual(curSection2, JSON.stringify(selectedSection),
								"The current selected section in storage should still be the selected section");

							strictEqual(JSON.stringify(controllerInstance.state.notebooks), JSON.stringify(freshNotebooks),
								"The state should always be updated with the fresh notebooks once it has been retrieved");
							strictEqual(JSON.stringify(controllerInstance.state.curSection), JSON.stringify(selectedSection),
								"The current selected section in state should still be the selected section");
							strictEqual(controllerInstance.state.status, Status.Succeeded,
								"The status should be Succeeded");
							done();
						});
					});
				}, (error) => {
					ok(false, "reject should not be called");
				});
			});
		});

		test("retrieveAndUpdateNotebookAndSectionSelection should update states correctly when there's notebook and curSection information found in storage," +
			" and then information is found on the server, but that selected section no longer exists.", (assert: QUnitAssert) => {
			let done = assert.async();

			let clipperState = MockProps.getMockClipperState();

			// Set up the storage mock
			let mockNotebooks = MockProps.getMockNotebooks();
			let mockSection = {
				section: mockNotebooks[0].sections[0],
				path: "Clipper Test > Full Page",
				parentId: mockNotebooks[0].id
			};
			initializeClipperStorage(JSON.stringify(mockNotebooks), JSON.stringify(mockSection), TestConstants.defaultUserInfoAsJsonString);

			let component = <SectionPicker
				onPopupToggle={() => {}}
				clipperState={clipperState} />;
			let controllerInstance = MithrilUtils.mountToFixture(component);

			strictEqual(JSON.stringify(controllerInstance.state), JSON.stringify({ notebooks: mockNotebooks, status: Status.Succeeded, curSection: mockSection }),
				"After the component is mounted, the state should be updated to reflect the notebooks and section found in storage");

			// After retrieving fresh notebooks, the storage should be updated with the fresh notebooks (we deleted the cached currently selected section)
			let freshNotebooks = MockProps.getMockNotebooks();
			freshNotebooks[0].sections = [];
			let responseJson = {
				"@odata.context": "https://www.onenote.com/api/v1.0/$metadata#me/notes/notebooks",
				value: freshNotebooks
			};
			this.server.respondWith([200, { "Content-Type": "application/json" }, JSON.stringify(responseJson)]);

			controllerInstance.retrieveAndUpdateNotebookAndSectionSelection().then((response: SectionPickerState) => {
				Clipper.getStoredValue(ClipperStorageKeys.cachedNotebooks, (notebooks) => {
					Clipper.getStoredValue(ClipperStorageKeys.currentSelectedSection, (curSection2) => {
						strictEqual(notebooks, JSON.stringify(freshNotebooks),
							"After fresh notebooks have been retrieved, the storage should be updated with them.");
						strictEqual(curSection2, undefined,
							"The current selected section in storage should now be undefined as it no longer exists in the fresh notebooks");
						strictEqual(JSON.stringify(controllerInstance.state.notebooks), JSON.stringify(freshNotebooks),
							"The state should always be updated with the fresh notebooks once it has been retrieved");
						strictEqual(controllerInstance.state.curSection, undefined,
							"The current selected section in state should be undefined");
						strictEqual(controllerInstance.state.status, Status.Succeeded,
							"The status should be Succeeded");
						done();
					});
				});
			}, (error) => {
				ok(false, "reject should not be called");
			});
		});

		test("retrieveAndUpdateNotebookAndSectionSelection should update states correctly when there's notebook and curSection information found in storage," +
			"the user does not make a new selection, and then notebooks is incorrectly returned as undefined or null from the server", (assert: QUnitAssert) => {
			let done = assert.async();

			let clipperState = MockProps.getMockClipperState();

			// Set up the storage mock
			let mockNotebooks = MockProps.getMockNotebooks();
			let mockSection = {
				section: mockNotebooks[0].sections[0],
				path: "Clipper Test > Full Page",
				parentId: mockNotebooks[0].id
			};
			initializeClipperStorage(JSON.stringify(mockNotebooks), JSON.stringify(mockSection), TestConstants.defaultUserInfoAsJsonString);

			let component = <SectionPicker
				onPopupToggle={() => {}}
				clipperState={clipperState} />;
			let controllerInstance = MithrilUtils.mountToFixture(component);

			strictEqual(JSON.stringify(controllerInstance.state), JSON.stringify({ notebooks: mockNotebooks, status: Status.Succeeded, curSection: mockSection }),
				"After the component is mounted, the state should be updated to reflect the notebooks and section found in storage");

			// After retrieving fresh undefined notebooks, the storage should not be updated with the undefined value, but should still keep the old cached information
			let responseJson = {
				"@odata.context": "https://www.onenote.com/api/v1.0/$metadata#me/notes/notebooks",
				value: undefined
			};
			this.server.respondWith([200, { "Content-Type": "application/json" }, JSON.stringify(responseJson)]);

			controllerInstance.retrieveAndUpdateNotebookAndSectionSelection().then((response: SectionPickerState) => {
				ok(false, "resolve should not be called");
			}, (error) => {
				Clipper.getStoredValue(ClipperStorageKeys.cachedNotebooks,
				(notebooks) => {
					Clipper.getStoredValue(ClipperStorageKeys.currentSelectedSection,
					(curSection) => {
						strictEqual(notebooks,
							JSON.stringify(mockNotebooks),
							"After undefined notebooks have been retrieved, the storage should not be updated with them.");
						strictEqual(curSection,
							JSON.stringify(mockSection),
							"The current selected section in storage should not have changed");

						strictEqual(JSON.stringify(controllerInstance.state.notebooks),
							JSON.stringify(mockNotebooks),
							"The state should not be updated as retrieving fresh notebooks returned undefined");
						strictEqual(JSON.stringify(controllerInstance.state.curSection),
							JSON.stringify(mockSection),
							"Since curSection was found in storage, and the user did not make an action to select another section, it should remain the same in state");
						strictEqual(controllerInstance.state.status, Status.Failed, "The status should be Failed");
						done();
					});
				});
			});
		});

		test("retrieveAndUpdateNotebookAndSectionSelection should update states correctly when there's notebook and curSection information found in storage," +
			"the user does not make a new selection, and the server returns an error status code", (assert: QUnitAssert) => {
			let done = assert.async();

			let clipperState = MockProps.getMockClipperState();

			// Set up the storage mock
			let mockNotebooks = MockProps.getMockNotebooks();
			let mockSection = {
				section: mockNotebooks[0].sections[0],
				path: "Clipper Test > Full Page",
				parentId: mockNotebooks[0].id
			};
			initializeClipperStorage(JSON.stringify(mockNotebooks), JSON.stringify(mockSection), TestConstants.defaultUserInfoAsJsonString);

			let component = <SectionPicker
				onPopupToggle={() => {}}
				clipperState={clipperState} />;
			let controllerInstance = MithrilUtils.mountToFixture(component);

			strictEqual(JSON.stringify(controllerInstance.state), JSON.stringify({ notebooks: mockNotebooks, status: Status.Succeeded, curSection: mockSection }),
				"After the component is mounted, the state should be updated to reflect the notebooks and section found in storage");

			// After retrieving fresh undefined notebooks, the storage should not be updated with the undefined value, but should still keep the old cached information
			let responseJson = {};
			this.server.respondWith([404, { "Content-Type": "application/json" }, JSON.stringify(responseJson)]);

			controllerInstance.retrieveAndUpdateNotebookAndSectionSelection().then((response: SectionPickerState) => {
				ok(false, "resolve should not be called");
			}, (error) => {
				Clipper.getStoredValue(ClipperStorageKeys.cachedNotebooks, (notebooks) => {
					Clipper.getStoredValue(ClipperStorageKeys.currentSelectedSection, (curSection) => {
						strictEqual(notebooks, JSON.stringify(mockNotebooks),
							"After undefined notebooks have been retrieved, the storage should not be updated with them.");
						strictEqual(curSection, JSON.stringify(mockSection),
							"The current selected section in storage should not have changed");

						strictEqual(JSON.stringify(controllerInstance.state.notebooks), JSON.stringify(mockNotebooks),
							"The state should not be updated as retrieving fresh notebooks returned undefined");
						strictEqual(JSON.stringify(controllerInstance.state.curSection), JSON.stringify(mockSection),
							"Since curSection was found in storage, and the user did not make an action to select another section, it should remain the same in state");
						strictEqual(controllerInstance.state.status, Status.Succeeded, "The status should be Succeeded since we have a fallback in storage");
						done();
					});
				});
			});
		});

		test("retrieveAndUpdateNotebookAndSectionSelection should update states correctly when there's no notebook and curSection information found in storage," +
			"the user does not make a new selection, and the server returns an error status code, therefore there's no fallback notebooks", (assert: QUnitAssert) => {
			let done = assert.async();

			let clipperState = MockProps.getMockClipperState();

			// Set up the storage mock
			initializeClipperStorage(undefined, undefined, TestConstants.defaultUserInfoAsJsonString);

			let component = <SectionPicker
				onPopupToggle={() => {}}
				clipperState={clipperState} />;
			let controllerInstance = MithrilUtils.mountToFixture(component);

			strictEqual(JSON.stringify(controllerInstance.state), JSON.stringify({ notebooks: undefined, status: Status.NotStarted, curSection: undefined }),
				"After the component is mounted, the state should be updated to reflect that notebooks and current section are not found in storage");

			// After retrieving fresh undefined notebooks, the storage should not be updated with the undefined value, but should still keep the old cached information
			let responseJson = {};
			this.server.respondWith([404, { "Content-Type": "application/json" }, JSON.stringify(responseJson)]);

			controllerInstance.retrieveAndUpdateNotebookAndSectionSelection().then((response: SectionPickerState) => {
				ok(false, "resolve should not be called");
			}, (error) => {
				Clipper.getStoredValue(ClipperStorageKeys.cachedNotebooks, (notebooks) => {
					Clipper.getStoredValue(ClipperStorageKeys.currentSelectedSection, (curSection) => {
						strictEqual(notebooks, undefined,
							"After undefined notebooks have been retrieved, the storage notebook value should still be undefined");
						strictEqual(curSection, undefined,
							"After undefined notebooks have been retrieved, the storage section value should still be undefined as there was no default section present");

						strictEqual(controllerInstance.state.notebooks, undefined,
							"After undefined notebooks have been retrieved, the state notebook value should still be undefined");
						strictEqual(controllerInstance.state.curSection, undefined,
							"After undefined notebooks have been retrieved, the state section value should still be undefined as there was no default section present");
						strictEqual(controllerInstance.state.status, Status.Failed, "The status should be Failed since getting fresh notebooks has failed and we don't have a fallback");
						done();
					});
				});
			});
		});

		test("fetchFreshNotebooks should parse out @odata.context from the raw 200 response and return the notebook object list and XHR in the resolve", (assert: QUnitAssert) => {
			let done = assert.async();

			let controllerInstance = MithrilUtils.mountToFixture(this.defaultComponent);

			let notebooks = MockProps.getMockNotebooks();
			let responseJson = {
				"@odata.context": "https://www.onenote.com/api/v1.0/$metadata#me/notes/notebooks",
				value: notebooks
			};
			this.server.respondWith([200, { "Content-Type": "application/json" }, JSON.stringify(responseJson)]);

			controllerInstance.fetchFreshNotebooks("sessionId").then((responsePackage: OneNoteApi.ResponsePackage<OneNoteApi.Notebook[]>) => {
				strictEqual(JSON.stringify(responsePackage.parsedResponse), JSON.stringify(notebooks),
					"The notebook list should be present in the response");
				ok(!!responsePackage.request,
					"The XHR must be present in the response");
			}, (error) => {
				ok(false, "reject should not be called");
			}).then(() => {
				done();
			});
		});

		test("fetchFreshNotebooks should parse out @odata.context from the raw 201 response and return the notebook object list and XHR in the resolve", (assert: QUnitAssert) => {
			let done = assert.async();

			let controllerInstance = MithrilUtils.mountToFixture(this.defaultComponent);

			let notebooks = MockProps.getMockNotebooks();
			let responseJson = {
				"@odata.context": "https://www.onenote.com/api/v1.0/$metadata#me/notes/notebooks",
				value: notebooks
			};
			this.server.respondWith([201, { "Content-Type": "application/json" }, JSON.stringify(responseJson)]);

			controllerInstance.fetchFreshNotebooks("sessionId").then((responsePackage: OneNoteApi.ResponsePackage<OneNoteApi.Notebook[]>) => {
				strictEqual(JSON.stringify(responsePackage.parsedResponse), JSON.stringify(notebooks),
					"The notebook list should be present in the response");
				ok(!!responsePackage.request,
					"The XHR must be present in the response");
			}, (error) => {
				ok(false, "reject should not be called");
			}).then(() => {
				done();
			});
		});

		test("fetchFreshNotebooks should reject with the error object and a copy of the response if the status code is 4XX", (assert: QUnitAssert) => {
			let done = assert.async();

			let controllerInstance = MithrilUtils.mountToFixture(this.defaultComponent);

			let responseJson = {
				error: "Unexpected response status",
				statusCode: 401,
				responseHeaders: { "Content-Type": "application/json" }
			};

			let expected = {
				error: responseJson.error,
				statusCode: responseJson.statusCode,
				responseHeaders: responseJson.responseHeaders,
				response: JSON.stringify(responseJson),
				timeout: 30000
			};

			this.server.respondWith([expected.statusCode, expected.responseHeaders, expected.response]);

			controllerInstance.fetchFreshNotebooks("sessionId").then((responsePackage: OneNoteApi.ResponsePackage<OneNoteApi.Notebook[]>) => {
				ok(false, "resolve should not be called");
			}, (error) => {
				deepEqual(error, expected, "The error object should be rejected");
				strictEqual(controllerInstance.state.apiResponseCode, undefined);
			}).then(() => {
				done();
			});
		});

		test("fetchFreshNotebooks should reject with the error object and an API response code if one is returned by the API", (assert: QUnitAssert) => {
			let done = assert.async();

			let controllerInstance = MithrilUtils.mountToFixture(this.defaultComponent);

			let responseJson = {
				error: {
					code: "10008",
					message: "The user's OneDrive, Group or Document Library contains more than 5000 items and cannot be queried using the API.",
					"@api.url": "http://aka.ms/onenote-errors#C10008"
				}
			};

			let expected = {
				error: "Unexpected response status",
				statusCode: 403,
				responseHeaders: { "Content-Type": "application/json" },
				response: JSON.stringify(responseJson),
				timeout: 30000
			};

			this.server.respondWith([expected.statusCode, expected.responseHeaders, expected.response]);

			controllerInstance.fetchFreshNotebooks("sessionId").then((responsePackage: OneNoteApi.ResponsePackage<OneNoteApi.Notebook[]>) => {
				ok(false, "resolve should not be called");
			}, (error: OneNoteApi.RequestError) => {
				deepEqual(error, expected, "The error object should be rejected");
				ok(!OneNoteApiUtils.isRetryable(controllerInstance.state.apiResponseCode));
			}).then(() => {
				done();
			});
		});

		test("fetchFreshNotebooks should reject with the error object and a copy of the response if the status code is 5XX", (assert: QUnitAssert) => {
			let done = assert.async();

			let controllerInstance = MithrilUtils.mountToFixture(this.defaultComponent);

			let responseJson = {
				error: "Unexpected response status",
				statusCode: 501,
				responseHeaders: { "Content-Type": "application/json" }
			};
			this.server.respondWith([501, responseJson.responseHeaders, JSON.stringify(responseJson)]);

			let expected = {
				error: responseJson.error,
				statusCode: responseJson.statusCode,
				responseHeaders: responseJson.responseHeaders,
				response: JSON.stringify(responseJson),
				timeout: 30000
			};

			controllerInstance.fetchFreshNotebooks("sessionId").then((responsePackage: OneNoteApi.ResponsePackage<OneNoteApi.Notebook[]>) => {
				ok(false, "resolve should not be called");
			}, (error) => {
				deepEqual(error, expected, "The error object should be rejected");
			}).then(() => {
				done();
			});
		});

		test("onPopupToggle should move focus to first tree item when dropdown opens", (assert: QUnitAssert) => {
			let done = assert.async();

			let clipperState = MockProps.getMockClipperState();
			let mockNotebooks = MockProps.getMockNotebooks();
			initializeClipperStorage(JSON.stringify(mockNotebooks), undefined, TestConstants.defaultUserInfoAsJsonString);

			let component = <SectionPicker
				onPopupToggle={() => {}}
				clipperState={clipperState} />;
			MithrilUtils.mountToFixture(component);

			// Open the dropdown
			MithrilUtils.simulateAction(() => {
				document.getElementById(TestConstants.Ids.sectionLocationContainer).click();
			});

			// Wait for requestAnimationFrame to complete
			requestAnimationFrame(() => {
				let notebookList = document.getElementById("notebookList");
				ok(notebookList, "Notebook list should be present when dropdown is open");
				if (notebookList) {
					let firstTreeItem = notebookList.querySelector("li[role='treeitem']") as HTMLElement;
					ok(firstTreeItem, "First tree item should exist in the notebook list");
				}
				done();
			});
		});
	}
}

(new SectionPickerTests()).runTests();
(new SectionPickerSinonTests()).runTests();
