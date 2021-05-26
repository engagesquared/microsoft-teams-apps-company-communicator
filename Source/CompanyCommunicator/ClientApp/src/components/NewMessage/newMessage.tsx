// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import * as React from 'react';
import { RouteComponentProps } from 'react-router-dom';
import { withTranslation, WithTranslation } from "react-i18next";
import { initializeIcons } from 'office-ui-fabric-react/lib/Icons';
import * as AdaptiveCards from "adaptivecards";
import { Button, Loader, Dropdown, Text, Flex, Input, TextArea, RadioGroup, DropdownProps, Slider } from '@fluentui/react-northstar'
import * as microsoftTeams from "@microsoft/teams-js";

import './newMessage.scss';
import './teamTheme.scss';
import { getDraftNotification, getTeams, createDraftNotification, updateDraftNotification, searchGroups, getGroups, verifyGroupAccess, updateNotification, getSentNotification } from '../../apis/messageListApi';
import {
    getInitAdaptiveCard, setCardTitle, setCardImageLink, setCardSummary,
    setCardAuthor, setCardBtn, setCardImageWidth, setCardImageHeight, setCardImageSize
} from '../AdaptiveCard/adaptiveCard';
import { getBaseUrl } from '../../configVariables';
import { ImageUtil } from '../../utility/imageutility';
import { TFunction } from "i18next";

type dropdownItem = {
    key: string,
    header: string,
    content: string,
    image: string,
    team: {
        id: string,
        teamGroupId: string,
    },
}

export interface IDraftMessage {
    id?: string,
    title: string,
    imageLink?: string,
    imageSize?: string,
    imageHeight?: number,
    imageWidth?: number,
    summary?: string,
    author: string,
    buttonTitle?: string,
    buttonLink?: string,
    teams: any[],
    rosters: any[],
    teamsGroups: any[],
    groups: any[],
    allUsers: boolean
}

export interface formState {
    title: string,
    summary?: string,
    btnLink?: string,
    imageLink?: string,
    imageSize?: string,
    imageWidth?: number,
    imageHeight?: number,
    btnTitle?: string,
    author: string,
    card?: any,
    page: string,
    teamsOptionSelected: boolean,
    rostersOptionSelected: boolean,
    allUsersOptionSelected: boolean,
    groupsOptionSelected: boolean,
    teams?: any[],
    groups?: any[],
    exists?: boolean,
    messageId: string,
    loader: boolean,
    groupAccess: boolean,
    loading: boolean,
    noResultMessage: string,
    unstablePinned?: boolean,
    selectedTeamsNum: number,
    selectedRostersNum: number,
    selectedGroupsNum: number,
    selectedRadioBtn: string,
    selectedTeams: dropdownItem[],
    selectedRosters: dropdownItem[],
    selectedGroups: dropdownItem[],
    errorImageUrlMessage: string,
    errorButtonUrlMessage: string,
    isEditSentMessage: boolean,
    submitIsLoading: boolean,
}

export interface INewMessageProps extends RouteComponentProps, WithTranslation {
    getDraftMessagesList?: any;
}

class NewMessage extends React.Component<INewMessageProps, formState> {
    readonly localize: TFunction;
    private card: any;
    private ratio: number = 1;
    private imageMaxWidth: number = 400;
    private imageMinWidth: number = 1;
    private imageMaxHeight: number = 400;
    private imageMinHeight: number = 1;

    constructor(props: INewMessageProps) {
        super(props);
        initializeIcons();
        this.localize = this.props.t;
        this.card = getInitAdaptiveCard(this.localize);
        this.setDefaultCard(this.card);

        this.state = {
            title: "",
            summary: "",
            author: "",
            btnLink: "",
            imageLink: "",
            btnTitle: "",
            card: this.card,
            page: "CardCreation",
            teamsOptionSelected: true,
            rostersOptionSelected: false,
            allUsersOptionSelected: false,
            groupsOptionSelected: false,
            messageId: "",
            loader: true,
            groupAccess: false,
            loading: false,
            noResultMessage: "",
            unstablePinned: true,
            selectedTeamsNum: 0,
            selectedRostersNum: 0,
            selectedGroupsNum: 0,
            selectedRadioBtn: "teams",
            selectedTeams: [],
            selectedRosters: [],
            selectedGroups: [],
            errorImageUrlMessage: "",
            errorButtonUrlMessage: "",
            isEditSentMessage: false,
            submitIsLoading: false,
        }
    }

    public async componentDidMount() {
        microsoftTeams.initialize();
        //- Handle the Esc key
        document.addEventListener("keydown", this.escFunction, false);
        let params = this.props.match.params;
        const isEditSentMessage = this.props.match.url.startsWith('/editmessage/');
        this.setGroupAccess();
        this.getTeamList().then(() => {
            if ('id' in params) {
                let id = params['id'];
                this.getItem(id, isEditSentMessage).then(() => {
                    const selectedTeams = this.makeDropdownItemList(this.state.selectedTeams || [], this.state.teams || []);
                    const selectedRosters = this.makeDropdownItemList(this.state.selectedRosters || [], this.state.teams || []);
                    this.setState({
                        exists: true,
                        messageId: id,
                        selectedTeams: selectedTeams,
                        selectedRosters: selectedRosters,
                        isEditSentMessage,
                    })
                });
                if (!isEditSentMessage) {
                    this.getGroupData(id).then(() => {
                        const selectedGroups = this.makeDropdownItems(this.state.groups);
                        this.setState({
                            selectedGroups: selectedGroups,
                            isEditSentMessage,
                        })
                    });
                }
            } else {
                this.setState({
                    exists: false,
                    loader: false
                }, () => {
                    let adaptiveCard = new AdaptiveCards.AdaptiveCard();
                    adaptiveCard.parse(this.state.card);
                    let renderedCard = adaptiveCard.render();
                    document.getElementsByClassName('adaptiveCardContainer')[0].appendChild(renderedCard);
                    if (this.state.btnLink) {
                        let link = this.state.btnLink;
                        adaptiveCard.onExecuteAction = function (action) { window.open(link, '_blank'); };
                    }
                })
            }
        });
    }

    private makeDropdownItems = (items: any[] | undefined) => {
        const resultedTeams: dropdownItem[] = [];
        if (items) {
            items.forEach((element) => {
                resultedTeams.push({
                    key: element.id,
                    header: element.name,
                    content: element.mail,
                    image: ImageUtil.makeInitialImage(element.name),
                    team: {
                        id: element.id,
                        teamGroupId: element.teamGroupId
                    },

                });
            });
        }
        return resultedTeams;
    }

    private makeDropdownItemList = (items: any[], fromItems: any[] | undefined) => {
        const dropdownItemList: dropdownItem[] = [];
        items.forEach(element =>
            dropdownItemList.push(
                typeof element !== "string" ? element : {
                    key: fromItems!.find(x => x.id === element).id,
                    header: fromItems!.find(x => x.id === element).name,
                    image: ImageUtil.makeInitialImage(fromItems!.find(x => x.id === element).name),
                    team: {
                        id: element
                    }
                })
        );
        return dropdownItemList;
    }

    public setDefaultCard = (card: any) => {
        const titleAsString = this.localize("TitleText");
        const summaryAsString = this.localize("Summary");
        const authorAsString = this.localize("Author1");
        const buttonTitleAsString = this.localize("ButtonTitle");

        setCardTitle(card, titleAsString);
        let imgUrl = getBaseUrl() + "/image/imagePlaceholder.png";
        setCardImageLink(card, imgUrl);
        setCardSummary(card, summaryAsString);
        setCardAuthor(card, authorAsString);
        setCardBtn(card, buttonTitleAsString, "https://adaptivecards.io");
    }

    private getTeamList = async () => {
        try {
            const response = await getTeams();
            this.setState({
                teams: response.data
            });
        } catch (error) {
            return error;
        }
    }

    private getGroupItems() {
        if (this.state.groups) {
            return this.makeDropdownItems(this.state.groups);
        }
        const dropdownItems: dropdownItem[] = [];
        return dropdownItems;
    }

    private setGroupAccess = async () => {
        await verifyGroupAccess().then(() => {
            this.setState({
                groupAccess: true
            });
        }).catch((error) => {
            const errorStatus = error.response.status;
            if (errorStatus === 403) {
                this.setState({
                    groupAccess: false
                });
            }
            else {
                throw error;
            }
        });
    }

    private getGroupData = async (id: number) => {
        try {
            const response = await getGroups(id);
            this.setState({
                groups: response.data
            });
        }
        catch (error) {
            return error;
        }
    }

    private getItem = async (id: number, isEditSentMessage: boolean) => {
        try {
            let response;
            if (isEditSentMessage) {
                response = await getSentNotification(id);
            } else {
                response = await getDraftNotification(id);
            }
            const draftMessageDetail = response.data;
            let selectedRadioButton = "teams";
            if (draftMessageDetail.rosters && draftMessageDetail.rosters.length > 0) {
                selectedRadioButton = "rosters";
            }
            else if (draftMessageDetail.groups && draftMessageDetail.groups.length > 0) {
                selectedRadioButton = "groups";
            }
            else if (draftMessageDetail.allUsers) {
                selectedRadioButton = "allUsers";
            }
            this.setState({
                teamsOptionSelected: draftMessageDetail.teams && draftMessageDetail.teams.length > 0,
                selectedTeamsNum: draftMessageDetail.teams && draftMessageDetail.teams.length,
                rostersOptionSelected: draftMessageDetail.rosters && draftMessageDetail.rosters.length > 0,
                selectedRostersNum: draftMessageDetail.rosters && draftMessageDetail.rosters.length,
                groupsOptionSelected: draftMessageDetail.groups && draftMessageDetail.groups.length > 0,
                selectedGroupsNum: draftMessageDetail.groups && draftMessageDetail.groups.length,
                selectedRadioBtn: selectedRadioButton,
                selectedTeams: draftMessageDetail.teams,
                selectedRosters: draftMessageDetail.rosters,
                selectedGroups: draftMessageDetail.groups
            });

            setCardTitle(this.card, draftMessageDetail.title);
            setCardImageLink(this.card, draftMessageDetail.imageLink);
            setCardImageSize(this.card, draftMessageDetail.imageSize);
            if (draftMessageDetail.imageSize === "Custom") {
                this.ratio = Number(draftMessageDetail.imageWidth / draftMessageDetail.imageHeight);
                setCardImageHeight(this.card, draftMessageDetail.imageHeight);
                setCardImageWidth(this.card, draftMessageDetail.imageWidth);
            }
            setCardSummary(this.card, draftMessageDetail.summary);
            setCardAuthor(this.card, draftMessageDetail.author);
            setCardBtn(this.card, draftMessageDetail.buttonTitle, draftMessageDetail.buttonLink);

            this.setState({
                title: draftMessageDetail.title,
                summary: draftMessageDetail.summary,
                btnLink: draftMessageDetail.buttonLink,
                imageLink: draftMessageDetail.imageLink,
                imageSize: draftMessageDetail.imageSize,
                imageHeight: draftMessageDetail.imageHeight,
                imageWidth: draftMessageDetail.imageWidth,
                btnTitle: draftMessageDetail.buttonTitle,
                author: draftMessageDetail.author,
                allUsersOptionSelected: draftMessageDetail.allUsers,
                loader: false
            }, () => {
                this.updateCard();
            });
        } catch (error) {
            return error;
        }
    }

    public componentWillUnmount() {
        document.removeEventListener("keydown", this.escFunction, false);
    }

    public render(): JSX.Element {
        if (this.state.loader) {
            return (
                <div className="Loader">
                    <Loader />
                </div>
            );
        } else {
            if (this.state.page === "CardCreation") {
                return (
                    <div className="taskModule">
                        <Flex column className="formContainer" vAlign="stretch" gap="gap.small" styles={{ background: "white" }}>
                            <Flex className="scrollableContent">
                                <Flex.Item size="size.half">
                                    <Flex column className="formContentContainer">
                                        <Input className="inputField"
                                            value={this.state.title}
                                            label={this.localize("TitleText")}
                                            placeholder={this.localize("PlaceHolderTitle")}
                                            onChange={this.onTitleChanged}
                                            autoComplete="off"
                                            fluid
                                        />

                                        <Input fluid className="inputField"
                                            value={this.state.imageLink}
                                            label={this.localize("ImageURL")}
                                            placeholder={this.localize("ImageURL")}
                                            onChange={this.onImageLinkChanged}
                                            error={!(this.state.errorImageUrlMessage === "")}
                                            autoComplete="off"
                                        />
                                        {this.state.imageLink && !this.state.errorImageUrlMessage ? (
                                            <>
                                                <div className="imageSizeContainer">
                                                    <Text>Image size</Text>
                                                    <Dropdown
                                                        items={["Auto", "Large", "Medium", "Small", "Custom"]}
                                                        defaultValue="Auto"
                                                        placeholder="Image size"
                                                        onChange={this.onImageSizeChanged}
                                                        value={this.state.imageSize}
                                                        checkable
                                                    />
                                                </div>
                                                {this.state.imageSize === "Custom" ? (
                                                    <>
                                                        <div className="imageSliderContainer">
                                                            <Text>Image width</Text>
                                                            <Flex vAlign="center" gap="gap.small">
                                                                <Slider
                                                                    min={this.imageMinWidth}
                                                                    max={this.imageMaxWidth}
                                                                    value={this.state.imageWidth}
                                                                    onChange={this.onImageWidthSliderChanged}
                                                                />
                                                                <Input
                                                                    type="number"
                                                                    input={{
                                                                        styles: {
                                                                            width: '90px',
                                                                        }
                                                                    }}
                                                                    value={this.state.imageWidth}
                                                                    icon={<span>px</span>}
                                                                    onChange={this.onImageWidthInputChanged}
                                                                />
                                                            </Flex>
                                                        </div>
                                                        <div className="imageSliderContainer">
                                                            <Text>Image height</Text>
                                                            <Flex vAlign="center" gap="gap.small">
                                                                <Slider
                                                                    min={this.imageMinHeight}
                                                                    max={this.imageMaxHeight}
                                                                    value={this.state.imageHeight}
                                                                    onChange={this.onImageHeightSliderChanged}
                                                                />
                                                                <Input
                                                                    type="number"
                                                                    input={{
                                                                        styles: {
                                                                            width: '90px',
                                                                        }
                                                                    }}
                                                                    value={this.state.imageHeight}
                                                                    icon={<span>px</span>}
                                                                    onChange={this.onImageHeightInputChanged}
                                                                />
                                                            </Flex>
                                                        </div>
                                                    </>
                                                ) : null}
                                            </>
                                        ) : null}
                                        <Text className={(this.state.errorImageUrlMessage === "") ? "hide" : "show"} error size="small" content={this.state.errorImageUrlMessage} />

                                        <div className="textArea">
                                            <Text content={this.localize("Summary")} />
                                            <TextArea
                                                autoFocus
                                                placeholder={this.localize("Summary")}
                                                value={this.state.summary}
                                                onChange={this.onSummaryChanged}
                                                fluid />
                                            <Flex vAlign="center" gap="gap.smaller">
                                                <Text size="small" as="div" disabled>
                                                    Supports basic Markdown expressions.
                                                </Text>
                                                <a href="https://docs.microsoft.com/en-us/microsoftteams/platform/task-modules-and-cards/cards/cards-format?tabs=adaptive-md%2Cconnector-html#formatting-cards-with-markdown" target="_blank" rel="noreferrer">
                                                    <Text size="small" as="div" disabled>Read more.</Text>
                                                </a>
                                            </Flex>
                                        </div>

                                        <Input className="inputField"
                                            value={this.state.author}
                                            label={this.localize("Author")}
                                            placeholder={this.localize("Author")}
                                            onChange={this.onAuthorChanged}
                                            autoComplete="off"
                                            fluid
                                        />
                                        <Input className="inputField"
                                            fluid
                                            value={this.state.btnTitle}
                                            label={this.localize("ButtonTitle")}
                                            placeholder={this.localize("ButtonTitle")}
                                            onChange={this.onBtnTitleChanged}
                                            autoComplete="off"
                                        />
                                        <Input className="inputField"
                                            fluid
                                            value={this.state.btnLink}
                                            label={this.localize("ButtonURL")}
                                            placeholder={this.localize("ButtonURL")}
                                            onChange={this.onBtnLinkChanged}
                                            error={!(this.state.errorButtonUrlMessage === "")}
                                            autoComplete="off"
                                        />
                                        <Text className={(this.state.errorButtonUrlMessage === "") ? "hide" : "show"} error size="small" content={this.state.errorButtonUrlMessage} />
                                    </Flex>
                                </Flex.Item>
                                <Flex.Item size="size.half">
                                    <div className="adaptiveCardContainer">
                                    </div>
                                </Flex.Item>
                            </Flex>

                            <Flex className="footerContainer" vAlign="end" hAlign="end">
                                <Flex className="buttonContainer">
                                    {this.state.isEditSentMessage ? 
                                     this.state.submitIsLoading ? <Loader size="small" /> : <Button content={"Save"} id="saveBtn" onClick={this.onSave} primary />
                                     : (
                                        <Button content={this.localize("Next")} disabled={this.isNextBtnDisabled()} id="saveBtn" onClick={this.onNext} primary />
                                    )}
                                </Flex>
                            </Flex>

                        </Flex>
                    </div>
                );
            }
            else if (this.state.page === "AudienceSelection") {
                return (
                    <div className="taskModule">
                        <Flex column className="formContainer" vAlign="stretch" gap="gap.small" styles={{ background: "white" }}>
                            <Flex className="scrollableContent">
                                <Flex.Item size="size.half">
                                    <Flex column className="formContentContainer">
                                        <h3>{this.localize("SendHeadingText")}</h3>
                                        <RadioGroup
                                            className="radioBtns"
                                            checkedValue={this.state.selectedRadioBtn}
                                            onCheckedValueChange={this.onGroupSelected}
                                            vertical={true}
                                            items={[
                                                {
                                                    name: "teams",
                                                    key: "teams",
                                                    value: "teams",
                                                    label: this.localize("SendToGeneralChannel"),
                                                    children: (Component, { name, ...props }) => {
                                                        return (
                                                            <Flex key={name} column>
                                                                <Component {...props} />
                                                                <Dropdown
                                                                    hidden={!this.state.teamsOptionSelected}
                                                                    placeholder={this.localize("SendToGeneralChannelPlaceHolder")}
                                                                    search
                                                                    multiple
                                                                    items={this.getItems()}
                                                                    value={this.state.selectedTeams}
                                                                    onChange={this.onTeamsChange}
                                                                    noResultsMessage={this.localize("NoMatchMessage")}
                                                                />
                                                            </Flex>
                                                        )
                                                    },
                                                },
                                                {
                                                    name: "rosters",
                                                    key: "rosters",
                                                    value: "rosters",
                                                    label: this.localize("SendToRosters"),
                                                    children: (Component, { name, ...props }) => {
                                                        return (
                                                            <Flex key={name} column>
                                                                <Component {...props} />
                                                                <Dropdown
                                                                    hidden={!this.state.rostersOptionSelected}
                                                                    placeholder={this.localize("SendToRostersPlaceHolder")}
                                                                    search
                                                                    multiple
                                                                    items={this.getItems()}
                                                                    value={this.state.selectedRosters}
                                                                    onChange={this.onRostersChange}
                                                                    unstable_pinned={this.state.unstablePinned}
                                                                    noResultsMessage={this.localize("NoMatchMessage")}
                                                                />
                                                            </Flex>
                                                        )
                                                    },
                                                },
                                                {
                                                    name: "allUsers",
                                                    key: "allUsers",
                                                    value: "allUsers",
                                                    label: this.localize("SendToAllUsers"),
                                                    children: (Component, { name, ...props }) => {
                                                        return (
                                                            <Flex key={name} column>
                                                                <Component {...props} />
                                                                <div className={this.state.selectedRadioBtn === "allUsers" ? "" : "hide"}>
                                                                    <div className="noteText">
                                                                        <Text error content={this.localize("SendToAllUsersNote")} />
                                                                    </div>
                                                                </div>
                                                            </Flex>
                                                        )
                                                    },
                                                },
                                                {
                                                    name: "groups",
                                                    key: "groups",
                                                    value: "groups",
                                                    label: this.localize("SendToGroups"),
                                                    children: (Component, { name, ...props }) => {
                                                        return (
                                                            <Flex key={name} column>
                                                                <Component {...props} />
                                                                <div className={this.state.groupsOptionSelected && !this.state.groupAccess ? "" : "hide"}>
                                                                    <div className="noteText">
                                                                        <Text error content={this.localize("SendToGroupsPermissionNote")} />
                                                                    </div>
                                                                </div>
                                                                <Dropdown
                                                                    className="hideToggle"
                                                                    hidden={!this.state.groupsOptionSelected || !this.state.groupAccess}
                                                                    placeholder={this.localize("SendToGroupsPlaceHolder")}
                                                                    search={this.onGroupSearch}
                                                                    multiple
                                                                    loading={this.state.loading}
                                                                    loadingMessage={this.localize("LoadingText")}
                                                                    items={this.getGroupItems()}
                                                                    value={this.state.selectedGroups}
                                                                    onSearchQueryChange={this.onGroupSearchQueryChange}
                                                                    onChange={this.onGroupsChange}
                                                                    noResultsMessage={this.state.noResultMessage}
                                                                    unstable_pinned={this.state.unstablePinned}
                                                                />
                                                                <div className={this.state.groupsOptionSelected && this.state.groupAccess ? "" : "hide"}>
                                                                    <div className="noteText">
                                                                        <Text error content={this.localize("SendToGroupsNote")} />
                                                                    </div>
                                                                </div>
                                                            </Flex>
                                                        )
                                                    },
                                                }
                                            ]}
                                        >

                                        </RadioGroup>
                                    </Flex>
                                </Flex.Item>
                                <Flex.Item size="size.half">
                                    <div className="adaptiveCardContainer">
                                    </div>
                                </Flex.Item>
                            </Flex>
                            <Flex className="footerContainer" vAlign="end" hAlign="end">
                                <Flex className="buttonContainer" gap="gap.small">
                                    <Flex.Item push>
                                        <Button content={this.localize("Back")} onClick={this.onBack} secondary />
                                    </Flex.Item>
                                    <Button content={this.localize("SaveAsDraft")} disabled={this.isSaveBtnDisabled()} id="saveBtn" onClick={this.onSave} primary />
                                </Flex>
                            </Flex>
                        </Flex>
                    </div>
                );
            } else {
                return (<div>Error</div>);
            }
        }
    }

    private onGroupSelected = (event: any, data: any) => {
        this.setState({
            selectedRadioBtn: data.value,
            teamsOptionSelected: data.value === 'teams',
            rostersOptionSelected: data.value === 'rosters',
            groupsOptionSelected: data.value === 'groups',
            allUsersOptionSelected: data.value === 'allUsers',
            selectedTeams: data.value === 'teams' ? this.state.selectedTeams : [],
            selectedTeamsNum: data.value === 'teams' ? this.state.selectedTeamsNum : 0,
            selectedRosters: data.value === 'rosters' ? this.state.selectedRosters : [],
            selectedRostersNum: data.value === 'rosters' ? this.state.selectedRostersNum : 0,
            selectedGroups: data.value === 'groups' ? this.state.selectedGroups : [],
            selectedGroupsNum: data.value === 'groups' ? this.state.selectedGroupsNum : 0,
        });
    }

    private isSaveBtnDisabled = () => {
        const teamsSelectionIsValid = (this.state.teamsOptionSelected && (this.state.selectedTeamsNum !== 0)) || (!this.state.teamsOptionSelected);
        const rostersSelectionIsValid = (this.state.rostersOptionSelected && (this.state.selectedRostersNum !== 0)) || (!this.state.rostersOptionSelected);
        const groupsSelectionIsValid = (this.state.groupsOptionSelected && (this.state.selectedGroupsNum !== 0)) || (!this.state.groupsOptionSelected);
        const nothingSelected = (!this.state.teamsOptionSelected) && (!this.state.rostersOptionSelected) && (!this.state.groupsOptionSelected) && (!this.state.allUsersOptionSelected);
        return (!teamsSelectionIsValid || !rostersSelectionIsValid || !groupsSelectionIsValid || nothingSelected)
    }

    private isNextBtnDisabled = () => {
        const title = this.state.title;
        const btnTitle = this.state.btnTitle;
        const btnLink = this.state.btnLink;
        return !(title && ((btnTitle && btnLink) || (!btnTitle && !btnLink)) && (this.state.errorImageUrlMessage === "") && (this.state.errorButtonUrlMessage === ""));
    }

    private getItems = () => {
        const resultedTeams: dropdownItem[] = [];
        if (this.state.teams) {
            let remainingUserTeams = this.state.teams;
            if (this.state.selectedRadioBtn !== "allUsers") {
                if (this.state.selectedRadioBtn === "teams") {
                    this.state.teams.filter(x => this.state.selectedTeams.findIndex(y => y.team.id === x.id) < 0);
                }
                else if (this.state.selectedRadioBtn === "rosters") {
                    this.state.teams.filter(x => this.state.selectedRosters.findIndex(y => y.team.id === x.id) < 0);
                }
            }
            remainingUserTeams.forEach((element) => {
                resultedTeams.push({
                    key: element.id,
                    header: element.name,
                    content: element.mail,
                    image: ImageUtil.makeInitialImage(element.name),
                    team: {
                        id: element.id,
                        teamGroupId: element.teamGroupId
                    }
                });
            });
        }
        return resultedTeams;
    }

    private static MAX_SELECTED_TEAMS_NUM: number = 20;

    private onTeamsChange = (event: any, itemsData: any) => {
        if (itemsData.value.length > NewMessage.MAX_SELECTED_TEAMS_NUM) return;
        this.setState({
            selectedTeams: itemsData.value,
            selectedTeamsNum: itemsData.value.length,
            selectedRosters: [],
            selectedRostersNum: 0,
            selectedGroups: [],
            selectedGroupsNum: 0
        })
    }

    private onRostersChange = (event: any, itemsData: any) => {
        if (itemsData.value.length > NewMessage.MAX_SELECTED_TEAMS_NUM) return;
        this.setState({
            selectedRosters: itemsData.value,
            selectedRostersNum: itemsData.value.length,
            selectedTeams: [],
            selectedTeamsNum: 0,
            selectedGroups: [],
            selectedGroupsNum: 0
        })
    }

    private onGroupsChange = (event: any, itemsData: any) => {
        this.setState({
            selectedGroups: itemsData.value,
            selectedGroupsNum: itemsData.value.length,
            groups: [],
            selectedTeams: [],
            selectedTeamsNum: 0,
            selectedRosters: [],
            selectedRostersNum: 0
        })
    }

    private onGroupSearch = (itemList: any, searchQuery: string) => {
        const result = itemList.filter(
            (item: { header: string; content: string; }) => (item.header && item.header.toLowerCase().indexOf(searchQuery.toLowerCase()) !== -1) ||
                (item.content && item.content.toLowerCase().indexOf(searchQuery.toLowerCase()) !== -1),
        )
        return result;
    }

    private onGroupSearchQueryChange = async (event: any, itemsData: any) => {

        if (!itemsData.searchQuery) {
            this.setState({
                groups: [],
                noResultMessage: "",
            });
        }
        else if (itemsData.searchQuery && itemsData.searchQuery.length <= 2) {
            this.setState({
                loading: false,
                noResultMessage: this.localize("NoMatchMessage"),
            });
        }
        else if (itemsData.searchQuery && itemsData.searchQuery.length > 2) {
            // handle event trigger on item select.
            const result = itemsData.items && itemsData.items.find(
                (item: { header: string; }) => item.header.toLowerCase() === itemsData.searchQuery.toLowerCase()
            )
            if (result) {
                return;
            }

            this.setState({
                loading: true,
                noResultMessage: "",
            });

            try {
                const query = encodeURIComponent(itemsData.searchQuery);
                const response = await searchGroups(query);
                this.setState({
                    groups: response.data,
                    loading: false,
                    noResultMessage: this.localize("NoMatchMessage")
                });
            }
            catch (error) {
                return error;
            }
        }
    }

    private onSave = () => {
        const selectedTeams: string[] = [];
        const selctedRosters: string[] = [];
        const selctedTeamGroups: string[] = [];
        const selectedGroups: string[] = [];
        if (this.state.selectedTeams) {
            this.state.selectedTeams.forEach(x => selectedTeams.push(x.team.id));
        }
        if (this.state.selectedRosters) {
            this.state.selectedRosters.forEach(x => {
                selctedRosters.push(x.team.id);
                selctedTeamGroups.push(x.team.teamGroupId);
            });
        }
        if (this.state.selectedGroups) {
            this.state.selectedGroups.forEach(x => selectedGroups.push(x.team.id));
        }

        const draftMessage: IDraftMessage = {
            id: this.state.messageId,
            title: this.state.title,
            imageLink: this.state.imageLink,
            imageSize: this.state.imageLink && !this.state.imageSize ? "Auto" : this.state.imageSize,
            imageHeight: this.state.imageSize === "Custom" ? this.state.imageHeight : undefined,
            imageWidth: this.state.imageSize === "Custom" ? this.state.imageWidth : undefined,
            summary: this.state.summary,
            author: this.state.author,
            buttonTitle: this.state.btnTitle,
            buttonLink: this.state.btnLink,
            teams: selectedTeams,
            rosters: selctedRosters,
            teamsGroups: selctedTeamGroups,
            groups: selectedGroups,
            allUsers: this.state.allUsersOptionSelected
        };

        if (this.state.isEditSentMessage) {
            this.setState({ submitIsLoading: true });
            this.editSentMessage(draftMessage).then(() => {
                this.setState({ submitIsLoading: false });
                microsoftTeams.tasks.submitTask();
            });
        } else if (this.state.exists) {
            this.editDraftMessage(draftMessage).then(() => {
                microsoftTeams.tasks.submitTask();
            });
        } else {
            this.postDraftMessage(draftMessage).then(() => {
                microsoftTeams.tasks.submitTask();
            });
        }
    }

    private editSentMessage = async (draftMessage: IDraftMessage) => {
        try {
            await updateNotification(draftMessage);
        } catch (error) {
            return error;
        }
    }

    private editDraftMessage = async (draftMessage: IDraftMessage) => {
        try {
            await updateDraftNotification(draftMessage);
        } catch (error) {
            return error;
        }
    }

    private postDraftMessage = async (draftMessage: IDraftMessage) => {
        try {
            await createDraftNotification(draftMessage);
        } catch (error) {
            throw error;
        }
    }

    public escFunction(event: any) {
        if (event.keyCode === 27 || (event.key === "Escape")) {
            microsoftTeams.tasks.submitTask();
        }
    }

    private onNext = (event: any) => {
        this.setState({
            page: "AudienceSelection"
        }, () => {
            this.updateCard();
        });
    }

    private onBack = (event: any) => {
        this.setState({
            page: "CardCreation"
        }, () => {
            this.updateCard();
        });
    }

    private onTitleChanged = (event: any) => {
        let showDefaultCard = (!event.target.value && !this.state.imageLink && !this.state.summary && !this.state.author && !this.state.btnTitle && !this.state.btnLink);
        setCardTitle(this.card, event.target.value);
        setCardImageLink(this.card, this.state.imageLink);
        setCardSummary(this.card, this.state.summary);
        setCardAuthor(this.card, this.state.author);
        setCardBtn(this.card, this.state.btnTitle, this.state.btnLink);
        this.setState({
            title: event.target.value,
            card: this.card
        }, () => {
            if (showDefaultCard) {
                this.setDefaultCard(this.card);
            }
            this.updateCard();
        });
    }

    private onImageLinkChanged = async (event: any) => {
        let url = event.target.value.toLowerCase();
        if (url === "") {
            this.setState({
                imageSize: ""
            });
            setCardImageSize(this.card, "");
        }
        if (!((url === "") || (url.startsWith("https://") || (url.startsWith("data:image/png;base64,")) || (url.startsWith("data:image/jpeg;base64,")) || (url.startsWith("data:image/gif;base64,"))))) {
            this.setState({
                imageSize: "",
                errorImageUrlMessage: this.localize("ErrorURLMessage")
            });
            setCardImageSize(this.card, "");
        } else {
            this.setState({
                errorImageUrlMessage: ""
            });
        }

        if (this.state.imageSize === "Custom") {
            const imageSize = await this.getImageSize(url || "");
            setCardImageWidth(this.card, imageSize.width);
            setCardImageHeight(this.card, imageSize.height);
            this.setState({
                imageWidth: imageSize.width,
                imageHeight: imageSize.height
            }, () => {
                this.updateCard();
            });
        }

        let showDefaultCard = (!this.state.title && !url && !this.state.summary && !this.state.author && !this.state.btnTitle && !this.state.btnLink);
        setCardTitle(this.card, this.state.title);
        setCardImageLink(this.card, url);
        setCardSummary(this.card, this.state.summary);
        setCardAuthor(this.card, this.state.author);
        setCardBtn(this.card, this.state.btnTitle, this.state.btnLink);
        this.setState({
            imageLink: url,
            card: this.card
        }, () => {
            if (showDefaultCard) {
                this.setDefaultCard(this.card);
            }
            this.updateCard();
        });
    }
    
    private getHeight(value: number) {
        var height = value / this.ratio;
        return Math.round(height);
    }

    private getWidth(value: number) {
        var width = this.ratio * value;
        return Math.round(width);
    }

    private getImageSize = async (url: string): Promise<any> => {
        return new Promise((res) => {
            var img = new Image();
            img.onload = () => {
                let width = img.width;
                let height = img.height;
                this.ratio = width / height;
                if (width > this.imageMaxWidth) {
                    width = this.imageMaxWidth;
                    height = this.getHeight(width);
                }
                if (height > this.imageMaxHeight) {
                    height = this.imageMaxHeight;
                    width = this.getWidth(height);
                }
                res({ width, height });
            };
            img.src = url;
        });
    }

    private onImageSizeChanged = async (event: React.MouseEvent | React.KeyboardEvent | null, data: DropdownProps) => {
        if (data.value === "Custom") {
            const imageSize = await this.getImageSize(this.state.imageLink || "");
            setCardImageWidth(this.card, imageSize.width);
            setCardImageHeight(this.card, imageSize.height);
            this.setState({
                imageWidth: imageSize.width,
                imageHeight: imageSize.height,
                imageSize: data.value
            }, () => {
                this.updateCard();
            });
            return;
        }

        setCardImageSize(this.card, data.value as string);
        this.setState({
            imageSize: data.value as string,
            card: this.card
        }, () => {
            this.updateCard();
        });
    }

    private onImageWidthChanged = (value: number) => {
        if (!this.state.imageWidth || !this.state.imageHeight) {
            return;
        }

        if(this.state.imageHeight <= this.imageMinHeight && value < this.state.imageWidth) {
            return;
        }

        if(this.state.imageHeight >= this.imageMaxHeight && value > this.state.imageWidth) {
            return;
        }

        let imageWidth = value < this.imageMaxWidth ? value : this.imageMaxWidth;
        let imageHeight = this.getHeight(imageWidth);
        if (imageHeight > this.imageMaxHeight) {
            imageHeight = this.imageMaxHeight;
            imageWidth = this.getWidth(imageHeight);
        }
        if (imageHeight < this.imageMinHeight) {
            imageHeight = this.imageMinHeight;
            imageWidth = this.getWidth(imageHeight);
        }
        setCardImageWidth(this.card, imageWidth);
        setCardImageHeight(this.card, imageHeight);
        this.setState({
            imageWidth: imageWidth,
            imageHeight: imageHeight,
            card: this.card
        }, () => {
            this.updateCard();
        });
    }

    private onImageWidthSliderChanged = (e: any, data: any) => {
        this.onImageWidthChanged(Number(data.value));
    }

    private onImageWidthInputChanged = (e: any) => {
        this.onImageWidthChanged(Number(e.target.value));
    }

    private onImageHeightChanged = (value: number) => {
        if (!this.state.imageWidth || !this.state.imageHeight) {
            return;
        }

        if(this.state.imageWidth <= this.imageMinWidth && value < this.state.imageHeight) {
            return;
        }

        if(this.state.imageWidth >= this.imageMaxWidth && value > this.state.imageHeight) {
            return;
        }

        let imageHeight = value < this.imageMaxHeight ? value : this.imageMaxHeight;
        let imageWidth = this.getWidth(imageHeight);
        if (imageWidth > this.imageMaxWidth) {
            imageWidth = this.imageMaxWidth;
            imageHeight = this.getHeight(imageWidth);
        }
        if (imageWidth < this.imageMinWidth) {
            imageWidth = this.imageMinWidth;
            imageHeight = this.getHeight(imageWidth);
        }
        setCardImageWidth(this.card, imageWidth);
        setCardImageHeight(this.card, imageHeight);
        this.setState({
            imageWidth: imageWidth,
            imageHeight: imageHeight,
            card: this.card
        }, () => {
            this.updateCard();
        });
    }

    private onImageHeightSliderChanged = (e: any, data: any) => {
        this.onImageHeightChanged(Number(data.value));
    }

    private onImageHeightInputChanged = (e: any) => {
        this.onImageHeightChanged(Number(e.target.value));
    }

    private onSummaryChanged = (event: any) => {
        let showDefaultCard = (!this.state.title && !this.state.imageLink && !event.target.value && !this.state.author && !this.state.btnTitle && !this.state.btnLink);
        setCardTitle(this.card, this.state.title);
        setCardImageLink(this.card, this.state.imageLink);
        setCardSummary(this.card, event.target.value);
        setCardAuthor(this.card, this.state.author);
        setCardBtn(this.card, this.state.btnTitle, this.state.btnLink);
        this.setState({
            summary: event.target.value,
            card: this.card
        }, () => {
            if (showDefaultCard) {
                this.setDefaultCard(this.card);
            }
            this.updateCard();
        });
    }

    private onAuthorChanged = (event: any) => {
        let showDefaultCard = (!this.state.title && !this.state.imageLink && !this.state.summary && !event.target.value && !this.state.btnTitle && !this.state.btnLink);
        setCardTitle(this.card, this.state.title);
        setCardImageLink(this.card, this.state.imageLink);
        setCardSummary(this.card, this.state.summary);
        setCardAuthor(this.card, event.target.value);
        setCardBtn(this.card, this.state.btnTitle, this.state.btnLink);
        this.setState({
            author: event.target.value,
            card: this.card
        }, () => {
            if (showDefaultCard) {
                this.setDefaultCard(this.card);
            }
            this.updateCard();
        });
    }

    private onBtnTitleChanged = (event: any) => {
        const showDefaultCard = (!this.state.title && !this.state.imageLink && !this.state.summary && !this.state.author && !event.target.value && !this.state.btnLink);
        setCardTitle(this.card, this.state.title);
        setCardImageLink(this.card, this.state.imageLink);
        setCardSummary(this.card, this.state.summary);
        setCardAuthor(this.card, this.state.author);
        if (event.target.value && this.state.btnLink) {
            setCardBtn(this.card, event.target.value, this.state.btnLink);
            this.setState({
                btnTitle: event.target.value,
                card: this.card
            }, () => {
                if (showDefaultCard) {
                    this.setDefaultCard(this.card);
                }
                this.updateCard();
            });
        } else {
            delete this.card.actions;
            this.setState({
                btnTitle: event.target.value,
            }, () => {
                if (showDefaultCard) {
                    this.setDefaultCard(this.card);
                }
                this.updateCard();
            });
        }
    }

    private onBtnLinkChanged = (event: any) => {
        if (!(event.target.value === "" || event.target.value.toLowerCase().startsWith("https://"))) {
            this.setState({
                errorButtonUrlMessage: this.localize("ErrorURLMessage")
            });
        } else {
            this.setState({
                errorButtonUrlMessage: ""
            });
        }

        const showDefaultCard = (!this.state.title && !this.state.imageLink && !this.state.summary && !this.state.author && !this.state.btnTitle && !event.target.value);
        setCardTitle(this.card, this.state.title);
        setCardSummary(this.card, this.state.summary);
        setCardAuthor(this.card, this.state.author);
        setCardImageLink(this.card, this.state.imageLink);
        if (this.state.btnTitle && event.target.value) {
            setCardBtn(this.card, this.state.btnTitle, event.target.value);
            this.setState({
                btnLink: event.target.value,
                card: this.card
            }, () => {
                if (showDefaultCard) {
                    this.setDefaultCard(this.card);
                }
                this.updateCard();
            });
        } else {
            delete this.card.actions;
            this.setState({
                btnLink: event.target.value
            }, () => {
                if (showDefaultCard) {
                    this.setDefaultCard(this.card);
                }
                this.updateCard();
            });
        }
    }

    private updateCard = () => {
        const adaptiveCard = new AdaptiveCards.AdaptiveCard();
        adaptiveCard.parse(this.state.card);
        const renderedCard = adaptiveCard.render();
        const container = document.getElementsByClassName('adaptiveCardContainer')[0].firstChild;
        if (container != null) {
            container.replaceWith(renderedCard);
        } else {
            document.getElementsByClassName('adaptiveCardContainer')[0].appendChild(renderedCard);
        }
        const link = this.state.btnLink;
        adaptiveCard.onExecuteAction = function (action) { window.open(link, '_blank'); }
    }
}

const newMessageWithTranslation = withTranslation()(NewMessage);
export default newMessageWithTranslation;
