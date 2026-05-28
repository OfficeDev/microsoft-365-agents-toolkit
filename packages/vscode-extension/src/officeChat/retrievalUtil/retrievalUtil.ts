// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { stemmer } from "./porter2Stemmer";
import stopwords from "../retrievalUtil/stop_words_english.json";

export type DocumentMetadata = {
  description: string;
  codeSample: string;
};

export type API = {
  name: string;
  description: string;
  kind: string;
  signature: string;
  examples: string[];
};

const synonymReplaceRules: Record<string, string> = {
  fetch: "get",
  retriev: "get",
  insert: "add",
  creat: "add",
  updat: "edit",
  modifi: "edit",
  remov: "delet",
};

// for new json array templates
export function prepareExamples(
  docs: DocumentMetadata[]
): [string[], Map<string, DocumentMetadata>] {
  const docsWithMetadata: Map<string, DocumentMetadata> = new Map();
  const cleanDocs: string[] = [];
  docs.forEach((doc) => {
    const cleanDescription = prepareDiscription(doc.description).join(" ");
    cleanDocs.push(cleanDescription);
    docsWithMetadata.set(cleanDescription, doc);
  });
  return [cleanDocs, docsWithMetadata];
}

export function filterStopWords(texts: string[]): string[] {
  return texts.filter((word) => !stopwords.includes(word));
}

export function keepLetters(text: string): string {
  return text.replace(/[^a-zA-Z ]/g, "");
}

export function stemText(texts: string[]): string[] {
  return texts.map(stemmer);
}

export function converseSynonym(text: string): string {
  return text in synonymReplaceRules ? synonymReplaceRules[text] : text;
}

export function stemAndSynonymConvese(texts: string[]): string[] {
  return texts.map(stemmer).map(converseSynonym);
}

export function prepareDiscription(text: string): string[] {
  return stemAndSynonymConvese(
    filterStopWords(
      keepLetters(text)
        .split(" ")
        .filter((word) => word.length > 0)
    )
  );
}
