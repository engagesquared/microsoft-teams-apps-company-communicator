// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import * as React from 'react';
import { withTranslation, WithTranslation } from "react-i18next";
import * as microsoftTeams from "@microsoft/teams-js";
import { Flex, Text, Loader } from '@fluentui/react-northstar';
import Message from './Message';
import { getHistoryNotifications } from '../../apis/messageListApi';
import { TFunction } from "i18next";
import './History.scss';

export interface IHistoryProps extends WithTranslation {
}

export interface IHistoryState {
    historyMessages: any[];
    isLoading: boolean;
}

class History extends React.Component<IHistoryProps, IHistoryState> {
    readonly localize: TFunction;
    constructor(props: IHistoryProps) {
        super(props);
        this.localize = this.props.t;
        this.state = {
            historyMessages: [],
            isLoading: true
        }
    }

    public async componentDidMount() {
        try {
            microsoftTeams.initialize();
            const res = await getHistoryNotifications();
            if (res && res.data) {
                const historyMessages = res.data.sort((a: any, b: any) => (new Date(a.createdDate).getTime()) - (new Date(b.createdDate).getTime()));
                this.setState({ historyMessages: historyMessages });
            }
        } finally {
            this.setState({ isLoading: false });
        }
    }

    public render(): JSX.Element {
        return (
            <Flex className="historyContainer" hAlign="center" vAlign="center" column fill gap="gap.small">
                {this.state.isLoading ? <Loader /> : null}
                {this.state.historyMessages && this.state.historyMessages.length ? (
                    this.state.historyMessages.map((message, idx) => <Message key={idx} message={message}/>)
                ) : !this.state.isLoading ? (
                    <Text>History is empty</Text>
                ) : null}
            </Flex>
        );
    }
}


const historyWithTranslation = withTranslation()(History);
export default historyWithTranslation;