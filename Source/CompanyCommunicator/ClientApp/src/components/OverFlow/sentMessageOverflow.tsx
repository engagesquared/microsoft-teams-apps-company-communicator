// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import React from 'react';
import { connect } from 'react-redux';
import { withTranslation, WithTranslation } from "react-i18next";
import { Menu, MoreIcon, Loader } from '@fluentui/react-northstar';
import { getBaseUrl } from '../../configVariables';
import * as microsoftTeams from "@microsoft/teams-js";
import { duplicateDraftNotification, deleteNotification } from '../../apis/messageListApi';
import { selectMessage, getMessagesList, getDraftMessagesList } from '../../actions';
import { TFunction } from "i18next";

export interface OverflowProps extends WithTranslation {
    message?: any;
    styles?: object;
    title?: string;
    selectMessage?: any;
    getMessagesList?: any;
    getDraftMessagesList?: any;
}

export interface OverflowState {
    menuOpen: boolean;
    isLoading: boolean;
}

export interface ITaskInfo {
    title?: string;
    height?: number;
    width?: number;
    url?: string;
    card?: string;
    fallbackUrl?: string;
    completionBotId?: string;
}

class Overflow extends React.Component<OverflowProps, OverflowState> {
    readonly localize: TFunction;
    constructor(props: OverflowProps) {
        super(props);
        this.localize = this.props.t;
        this.state = {
            menuOpen: false,
            isLoading: false,
        };
    }

    public componentDidMount() {
        microsoftTeams.initialize();
    }

    public render(): JSX.Element {
        const items = [
            {
                key: 'more',
                icon: this.state.isLoading ? <Loader size="small" /> : <MoreIcon outline={true} />,
                menuOpen: this.state.menuOpen,
                active: this.state.menuOpen,
                indicator: false,
                menu: {
                    items: [
                        {
                            key: 'status',
                            content: this.localize("ViewStatus"),
                            onClick: (event: any) => {
                                event.stopPropagation();
                                this.setState({
                                    menuOpen: false,
                                });
                                let url = getBaseUrl() + "/viewstatus/" + this.props.message.id + "?locale={locale}";
                                this.onOpenTaskModule(null, url, this.localize("ViewStatus"));
                            }
                        },
                        {
                            key: 'edit',
                            content: this.localize("Edit"),
                            onClick: (event: any) => {
                                event.stopPropagation();
                                this.setState({
                                    menuOpen: false,
                                });
                                let url = getBaseUrl() + "/editmessage/" + this.props.message.id + "?locale={locale}";
                                this.onOpenTaskModule(null, url, this.localize("Edit"));
                            }
                        },
                        {
                            key: 'delete',
                            content: this.localize("Delete"),
                            onClick: async (event: any) => {
                                event.stopPropagation();
                                this.setState({
                                    menuOpen: false,
                                    isLoading: true,
                                });
                                await deleteNotification(this.props.message.id);
                                await this.props.getMessagesList();
                                this.setState({
                                    isLoading: false,
                                });
                            }
                        },
                        {
                            key: 'duplicate',
                            content: this.localize("Duplicate"),
                            onClick: (event: any) => {
                                event.stopPropagation();
                                this.setState({
                                    menuOpen: false,
                                });
                                this.duplicateDraftMessage(this.props.message.id).then(() => {
                                    this.props.getDraftMessagesList();
                                });
                            }
                        },
                    ],
                },
                onMenuOpenChange: (e: any, { menuOpen }: any) => {
                    this.setState({
                        menuOpen: menuOpen
                    });
                },
            },
        ];

        return <Menu className="menuContainer" iconOnly items={items} styles={this.props.styles} title={this.props.title} />;
    }

    private onOpenTaskModule = (event: any, url: string, title: string) => {
        let taskInfo: ITaskInfo = {
            url: url,
            title: title,
            height: 530,
            width: 1000,
            fallbackUrl: url,
        };
        let submitHandler = (err: any, result: any) => {
            this.props.getMessagesList();
        };
        microsoftTeams.tasks.startTask(taskInfo, submitHandler);
    }

    private duplicateDraftMessage = async (id: number) => {
        try {
            await duplicateDraftNotification(id);
        } catch (error) {
            return error;
        }
    }
}

const mapStateToProps = (state: any) => {
    return { messagesList: state.messagesList };
}

const overflowWithTranslation = withTranslation()(Overflow);
export default connect(mapStateToProps, { selectMessage, getMessagesList, getDraftMessagesList })(overflowWithTranslation);
