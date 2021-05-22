// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import * as React from 'react';
import { withTranslation, WithTranslation } from "react-i18next";
import { Text, Flex, Ref, Segment } from '@fluentui/react-northstar';
import * as AdaptiveCards from "adaptivecards";
import {
    getInitAdaptiveCard, setCardTitle, setCardImageLink, setCardSummary,
    setCardAuthor, setCardBtn, setCardImageWidth, setCardImageHeight, setCardImageSize
} from '../AdaptiveCard/adaptiveCard';
import { TFunction } from "i18next";
import './Message.scss';

export interface IMessageProps extends WithTranslation {
    message: any;
}

class Message extends React.Component<IMessageProps> {
    readonly localize: TFunction;
    private ref: any;
    constructor(props: IMessageProps) {
        super(props);
        this.localize = this.props.t;
        this.ref = React.createRef();
    }

    public async componentDidMount() {
        this.renderCard()
    }

    public render(): JSX.Element {
        const createdString = this.props.message.createdDate && new Date(this.props.message.createdDate).toLocaleString();
        return (
            <Segment className="segment">
                <Flex className="historyContainer" column>
                    <Flex gap="gap.small">
                        <Text as="div" weight="bold" size="small">Company Communicator</Text>
                        <Text as="div" size="small">{createdString}</Text>
                    </Flex>
                    <Ref innerRef={this.ref}>
                        <div />
                    </Ref>
                </Flex>
            </Segment>
        );
    }

    private renderCard = () => {
        const historyMessage = this.props.message;
        if (!historyMessage) {
            return;
        }

        const container = this.ref.current;
        if (!container) {
            return;
        }
        
        const adaptiveCard = new AdaptiveCards.AdaptiveCard();
        const card = getInitAdaptiveCard(this.localize);
        setCardTitle(card, historyMessage.title);
        setCardSummary(card, historyMessage.summary);
        setCardImageLink(card, historyMessage.imageLink);
        setCardImageSize(card, historyMessage.imageSize);
        if (historyMessage.imageSize === "Custom") {
            setCardImageHeight(card, historyMessage.imageHeight);
            setCardImageWidth(card, historyMessage.imageWidth);
        }
        setCardAuthor(card, historyMessage.author);
        setCardBtn(card, historyMessage.buttonTitle, historyMessage.buttonLink);
        adaptiveCard.onExecuteAction = function (action) { window.open(historyMessage.buttonLink, '_blank'); }
        adaptiveCard.parse(card);
        const renderedCard = adaptiveCard.render();
        container.appendChild(renderedCard);
    }
}


const messageWithTranslation = withTranslation()(Message);
export default messageWithTranslation;