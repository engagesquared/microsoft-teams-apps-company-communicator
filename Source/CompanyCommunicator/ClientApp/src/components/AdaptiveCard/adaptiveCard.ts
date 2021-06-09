// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import * as AdaptiveCards from "adaptivecards";
import MarkdownIt from "markdown-it";
import { TFunction } from "i18next";

AdaptiveCards.AdaptiveCard.onProcessMarkdown = function(text, result) {
    const md = new MarkdownIt();
    // 'blockquote', 'reference', 'paragraph' currently don't supported by adaptive cards, but disabling these rules causes infinity loop
    md.block.ruler.enableOnly(['list', 'blockquote', 'reference', 'paragraph']);
    md.inline.ruler.enableOnly(['text', 'emphasis', 'link']);
	result.outputHtml = md.render(text);
	result.didProcess = true;
}

export const getInitAdaptiveCard = (t: TFunction) => {
    const titleTextAsString = t("TitleText");
    return (
        {
            "type": "AdaptiveCard",
            "body": [
                {
                    "type": "TextBlock",
                    "text": "IMPORTANT!",
                    "wrap": true,
                    "size": "medium",
                    "color": "attention",
                    "weight": "bolder",
                    "isVisible": false
                },
                {
                    "type": "TextBlock",
                    "weight": "Bolder",
                    "text": titleTextAsString,
                    "size": "ExtraLarge",
                    "wrap": true
                },
                {
                    "type": "Image",
                    "spacing": "Default",
                    "url": "",
                    "size": "Auto",
                    "width": "",
                    "height": "",
                    "altText": ""
                },
                {
                    "type": "TextBlock",
                    "text": "",
                    "wrap": true
                },
                {
                    "type": "TextBlock",
                    "wrap": true,
                    "size": "Small",
                    "weight": "Lighter",
                    "text": ""
                }
            ],
            "$schema": "https://adaptivecards.io/schemas/adaptive-card.json",
            "version": "1.0"
        }
    );
}

export const setCardImportant = (card: any, important: boolean) => {
    card.body[0].isVisible = important;
}

export const getCardTitle = (card: any) => {
    return card.body[1].text;
}

export const setCardTitle = (card: any, title: string) => {
    card.body[1].text = title;
}

export const getCardImageLink = (card: any) => {
    return card.body[2].url;
}

export const setCardImageLink = (card: any, imageLink?: string) => {
    card.body[2].url = imageLink;
}

export const setCardImageWidth = (card: any, width: number) => {
    card.body[2].width = `${width}px`;
    card.body[2].size = null;
}

export const setCardImageHeight = (card: any, height: number) => {
    card.body[2].height = `${height}px`;
    card.body[2].size = null;
}

export const setCardImageSize = (card: any, size: string) => {
    card.body[2].size = size;
    card.body[2].width = null;
    card.body[2].height = null;
}

export const getCardSummary = (card: any) => {
    return card.body[3].text;
}

export const setCardSummary = (card: any, summary?: string) => {
    card.body[3].text = summary;
}

export const getCardAuthor = (card: any) => {
    return card.body[4].text;
}

export const setCardAuthor = (card: any, author?: string) => {
    card.body[4].text = author;
}

export const getCardBtnTitle = (card: any) => {
    return card.actions[0].title;
}

export const getCardBtnLink = (card: any) => {
    return card.actions[0].url;
}

export const setCardBtn = (card: any, buttonTitle?: string, buttonLink?: string) => {
    if (buttonTitle && buttonLink) {
        card.actions = [
            {
                "type": "Action.OpenUrl",
                "title": buttonTitle,
                "url": buttonLink
            }
        ];
    } else {
        delete card.actions;
    }
}
