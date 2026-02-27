import {Constants} from "../../constants";
import {AugmentationHelper} from "../../contentCapture/augmentationHelper";
import {ExtensionUtils} from "../../extensions/extensionUtils";
import {InvokeMode} from "../../extensions/invokeOptions";
import {Localization} from "../../localization/localization";
import {ClipMode} from "../clipMode";
import {ClipperStateProp} from "../clipperState";
import {ComponentBase} from "../componentBase";
import {ModeButton, PropsForModeElementNoAriaGrouping} from "./modeButton";

class ModeButtonSelectorClass extends ComponentBase<{}, ClipperStateProp> {
	onModeSelected(newMode: ClipMode) {
		this.props.clipperState.setState({
			currentMode: this.props.clipperState.currentMode.set(newMode),
			focusOnRender: Constants.Ids.previewInnerContainer
		});
	};

	private getPdfButtonProps(currentMode: ClipMode, tabIndex: number): PropsForModeElementNoAriaGrouping {
		if (this.props.clipperState.pageInfo.contentType !== OneNoteApi.ContentType.EnhancedUrl) {
			return undefined;
		}

		return {
			imgSrc: ExtensionUtils.getImageResourceUrl("pdf.png"),
			label: Localization.getLocalizedString("WebClipper.ClipType.Pdf.Button"),
			myMode: ClipMode.Pdf,
			selected: currentMode === ClipMode.Pdf,
			onModeSelected: this.onModeSelected.bind(this),
			tooltipText: Localization.getLocalizedString("WebClipper.ClipType.Pdf.Button.Tooltip"),
			tabIndex: tabIndex
		};
	}

	private getAugmentationButtonProps(currentMode: ClipMode, tabIndex: number): PropsForModeElementNoAriaGrouping {
		if (this.props.clipperState.pageInfo.contentType === OneNoteApi.ContentType.EnhancedUrl) {
			return undefined;
		}

		let augmentationType: string = AugmentationHelper.getAugmentationType(this.props.clipperState);
		let augmentationLabel: string = Localization.getLocalizedString("WebClipper.ClipType." + augmentationType + ".Button");
		let augmentationTooltip = Localization.getLocalizedString("WebClipper.ClipType.Button.Tooltip").replace("{0}", augmentationLabel);
		let buttonSelected: boolean = currentMode === ClipMode.Augmentation;
		return {
			imgSrc: (augmentationType === "Article") ? ExtensionUtils.getImageResourceUrl("article.svg") : ExtensionUtils.getImageResourceUrl(augmentationType + ".png"),
			label: augmentationLabel,
			myMode: ClipMode.Augmentation,
			selected: buttonSelected,
			onModeSelected: this.onModeSelected.bind(this),
			tooltipText: augmentationTooltip,
			tabIndex: tabIndex
		};
	}

	private getFullPageButtonProps(currentMode: ClipMode, tabIndex: number): PropsForModeElementNoAriaGrouping {
		if (this.props.clipperState.pageInfo.contentType === OneNoteApi.ContentType.EnhancedUrl) {
			return undefined;
		}

		return {
			imgSrc: ExtensionUtils.getImageResourceUrl("fullpage.svg"),
			label: Localization.getLocalizedString("WebClipper.ClipType.ScreenShot.Button"),
			myMode: ClipMode.FullPage,
			selected: currentMode === ClipMode.FullPage,
			onModeSelected: this.onModeSelected.bind(this),
			tooltipText: Localization.getLocalizedString("WebClipper.ClipType.ScreenShot.Button.Tooltip"),
			tabIndex: tabIndex
		};
	}

	private getRegionButtonProps(currentMode: ClipMode, tabIndex: number): PropsForModeElementNoAriaGrouping {
		let enableRegionClipping = this.props.clipperState.injectOptions && this.props.clipperState.injectOptions.enableRegionClipping;
		let contextImageModeUsed = this.props.clipperState.invokeOptions && this.props.clipperState.invokeOptions.invokeMode === InvokeMode.ContextImage;

		if (!enableRegionClipping && !contextImageModeUsed) {
			return undefined;
		}

		return {
			imgSrc: ExtensionUtils.getImageResourceUrl("region.svg"),
			label: Localization.getLocalizedString(this.getRegionButtonLabel()),
			myMode: ClipMode.Region,
			selected: currentMode === ClipMode.Region,
			onModeSelected: this.onModeSelected.bind(this),
			tooltipText: Localization.getLocalizedString("WebClipper.ClipType.MultipleRegions.Button.Tooltip"),
			tabIndex: tabIndex
		};
	}

	private getRegionButtonLabel(): string {
		return "WebClipper.ClipType.Region.Button";
	}

	private getSelectionButtonProps(currentMode: ClipMode, tabIndex: number): PropsForModeElementNoAriaGrouping {
		if (this.props.clipperState.invokeOptions.invokeMode !== InvokeMode.ContextTextSelection) {
			return undefined;
		}

		return {
			imgSrc: ExtensionUtils.getImageResourceUrl("select.png"),
			label: Localization.getLocalizedString("WebClipper.ClipType.Selection.Button"),
			myMode: ClipMode.Selection,
			selected: currentMode === ClipMode.Selection,
			onModeSelected: this.onModeSelected.bind(this),
			tooltipText: Localization.getLocalizedString("WebClipper.ClipType.Selection.Button.Tooltip"),
			tabIndex: tabIndex
		};
	}

	private getBookmarkButtonProps(currentMode: ClipMode, tabIndex: number): PropsForModeElementNoAriaGrouping {
		if (this.props.clipperState.pageInfo.rawUrl.indexOf("file:///") === 0) {
			return undefined;
		}

		return {
			imgSrc: ExtensionUtils.getImageResourceUrl("bookmark.svg"),
			label: Localization.getLocalizedString("WebClipper.ClipType.Bookmark.Button"),
			myMode: ClipMode.Bookmark,
			selected: currentMode === ClipMode.Bookmark,
			onModeSelected: this.onModeSelected.bind(this),
			tooltipText: Localization.getLocalizedString("WebClipper.ClipType.Bookmark.Button.Tooltip"),
			tabIndex: tabIndex
		};
	}

	private getListOfButtons(): HTMLElement[] {
		let currentMode = this.props.clipperState.currentMode.get();

		// Base tabIndex for mode buttons - they should come before PDF options (60+) and location dropdown (70)
		let baseTabIndex = 40;

		let buttonProps = [
			this.getFullPageButtonProps(currentMode, baseTabIndex),
			this.getRegionButtonProps(currentMode, baseTabIndex + 1),
			this.getAugmentationButtonProps(currentMode, baseTabIndex + 2),
			this.getSelectionButtonProps(currentMode, baseTabIndex + 3),
			this.getBookmarkButtonProps(currentMode, baseTabIndex + 4),
			this.getPdfButtonProps(currentMode, baseTabIndex + 5),
		];

		let visibleButtons = [];

		let propsForVisibleButtons = buttonProps.filter(attributes => !!attributes);
		for (let i = 0; i < propsForVisibleButtons.length; i++) {
			let attributes = propsForVisibleButtons[i];
			let ariaPos = i + 1;
			visibleButtons.push(<ModeButton {...attributes} aria-setsize={propsForVisibleButtons.length}
				aria-posinset={ariaPos} aria-selected={attributes.selected} clipperState={this.props.clipperState} />);
		}
		return visibleButtons;
	}

	public render() {
		let currentMode = this.props.clipperState.currentMode.get();

		return (
			<div role="group">
				<div style={Localization.getFontFamilyAsStyle(Localization.FontFamily.Semilight)}
					role="listbox" className="modeButtonContainer">
					{ this.getListOfButtons() }
				</div>
			</div>
		);
	}
}

let component = ModeButtonSelectorClass.componentize();
export {component as ModeButtonSelector};
