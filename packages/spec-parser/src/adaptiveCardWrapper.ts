// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
"use strict";

import { ResponseSemanticsObject } from "@microsoft/app-manifest";
import { ConstantString } from "./constants";
import {
  AdaptiveCard,
  ArrayElement,
  ImageElement,
  InferredProperties,
  PreviewCardTemplate,
  TextBlockElement,
  WrappedAdaptiveCard,
} from "./interfaces";
import { Utils } from "./utils";

export function wrapAdaptiveCard(card: AdaptiveCard, jsonPath: string): WrappedAdaptiveCard {
  const result: WrappedAdaptiveCard = {
    version: ConstantString.WrappedCardVersion,
    $schema: ConstantString.WrappedCardSchema,
    jsonPath: jsonPath,
    responseLayout: ConstantString.WrappedCardResponseLayout,
    responseCardTemplate: card,
    previewCardTemplate: inferPreviewCardTemplate(card),
  };

  return result;
}

export function wrapResponseSemantics(
  card: AdaptiveCard,
  jsonPath: string
): ResponseSemanticsObject {
  const props = inferProperties(card);
  const dataPath = jsonPath === "$" ? "$" : "$." + jsonPath;
  const result: ResponseSemanticsObject = {
    data_path: dataPath,
  };

  if (props.title || props.subtitle || props.imageUrl) {
    result.properties = {};
    if (props.title) {
      result.properties.title = "$." + props.title;
    }

    if (props.subtitle) {
      result.properties.subtitle = "$." + props.subtitle;
    }

    if (props.imageUrl) {
      result.properties.url = "$." + props.imageUrl;
    }
  }

  result.static_template = card as any;

  return result;
}

/**
 * Infers the preview card template from an Adaptive Card and a JSON path.
 * The preview card template includes a title and an optional subtitle and image.
 * It populates the preview card template with the first text block that matches
 * each well-known name, in the order of title, subtitle, and image.
 * If no text block matches the title or subtitle, it uses the first two text block as the title and subtitle.
 * If the title is still empty and the subtitle is not empty, it uses subtitle as the title.
 * @param card The Adaptive Card to infer the preview card template from.
 * @param jsonPath The JSON path to the root object in the card body.
 * @returns The inferred preview card template.
 */
export function inferPreviewCardTemplate(card: AdaptiveCard): PreviewCardTemplate {
  const result: PreviewCardTemplate = {
    title: "result",
  };
  const inferredProperties = inferProperties(card);
  if (inferredProperties.title) {
    result.title = `\${if(${inferredProperties.title}, ${inferredProperties.title}, 'N/A')}`;
  }

  if (inferredProperties.subtitle) {
    result.subtitle = `\${if(${inferredProperties.subtitle}, ${inferredProperties.subtitle}, 'N/A')}`;
  }

  if (inferredProperties.imageUrl) {
    result.image = {
      url: `\${${inferredProperties.imageUrl}}`,
      alt: `\${if(${inferredProperties.imageUrl}, ${inferredProperties.imageUrl}, 'N/A')}`,
      $when: `\${${inferredProperties.imageUrl} != null && ${inferredProperties.imageUrl} != ''}`,
    };
  }

  return result;
}

function inferProperties(card: AdaptiveCard): InferredProperties {
  const result: InferredProperties = {};

  const nameSet = new Set<string>();

  let rootObject: (TextBlockElement | ArrayElement | ImageElement)[];
  if (card.body[0]?.type === ConstantString.ContainerType) {
    rootObject = (card.body[0] as ArrayElement).items;
  } else {
    rootObject = card.body;
  }

  for (const element of rootObject) {
    if (element.type === ConstantString.TextBlockType) {
      const textElement = element as TextBlockElement;
      const index = textElement.text.indexOf("${if(");
      if (index > 0) {
        const text = textElement.text.substring(index);
        const match = text.match(/\${if\(([^,]+),/);
        const property = match ? match[1] : "";
        if (property) {
          nameSet.add(property);
        }
      }
    } else if (element.type === ConstantString.ImageType) {
      const imageElement = element as ImageElement;
      const match = imageElement.url.match(/\${([^,]+)}/);
      const property = match ? match[1] : "";
      if (property) {
        nameSet.add(property);
      }
    }
  }

  for (const name of nameSet) {
    if (!result.title && Utils.isWellKnownName(name, ConstantString.WellknownTitleName)) {
      result.title = name;
      nameSet.delete(name);
    } else if (
      !result.subtitle &&
      Utils.isWellKnownName(name, ConstantString.WellknownSubtitleName)
    ) {
      result.subtitle = name;
      nameSet.delete(name);
    } else if (!result.imageUrl && Utils.isWellKnownName(name, ConstantString.WellknownImageName)) {
      result.imageUrl = name;
      nameSet.delete(name);
    }
  }

  for (const name of nameSet) {
    if (!result.title) {
      result.title = name;
      nameSet.delete(name);
    } else if (!result.subtitle) {
      result.subtitle = name;
      nameSet.delete(name);
    }
  }

  if (!result.title && result.subtitle) {
    result.title = result.subtitle;
    delete result.subtitle;
  }

  return result;
}
