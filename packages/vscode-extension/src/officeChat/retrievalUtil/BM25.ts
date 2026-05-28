// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

export type DocumentWithmetadata = {
  documentText: string;
  metadata: object | null;
};

export type BMDocument = {
  score: number;
  document: DocumentWithmetadata;
};

export type BM25Config = {
  //Term frequency saturation parameter. Recommended value: between 1.2 and 2
  b?: number;
  //Length normalization parameter. Recommended value: > 0.75
  k1?: number;
  // Frequency normalization lower bound. Recommended value: between 0.5 and 1
  d?: number;
  // Frequency in the query weight. Default is 1, 0 means do not consider compeated terms in the query
  k3?: number;
};

export class BM25 {
  b: number;
  k1: number;
  d: number;
  k3: number;
  averageLength: number;
  documents: DocumentWithmetadata[];

  constructor(documents: DocumentWithmetadata[], config?: BM25Config) {
    this.b = config && config.b ? config.b : 0.75;
    this.k1 = config && config.k1 ? config.k1 : 1.2;
    this.d = config && config.d ? config.d : 0;
    this.k3 = config && config.k3 ? config.k3 : 1;

    this.documents = documents;
    this.averageLength =
      this.documents.reduce((acc, doc) => acc + this.countWords(doc.documentText), 0) /
      this.documents.length;
  }

  private countWords(s: string): number {
    const matches = s.match(/\b[\w']+\b/g);
    return matches ? matches.length : 0;
  }

  private countFrequency(word: string, singleDocuemnt: string): number {
    const regex = new RegExp(`\\b${word}\\b`, "g");
    return singleDocuemnt.match(regex)?.length || 0;
  }

  private getIDF(word: string): number {
    const docCount = this.documents.length;
    const relevantDocCount = this.documents.filter(
      (doc) => this.countFrequency(word, doc.documentText) > 0
    ).length;
    return Math.log(1 + (docCount - relevantDocCount + 0.5) / (relevantDocCount + 0.5));
  }

  private score(word: string, singleDocuemnt: string): number {
    const frequency = this.countFrequency(word, singleDocuemnt);
    return (
      this.getIDF(word) *
      (this.d +
        ((this.k1 + 1) * frequency) /
          (frequency +
            this.k1 *
              (1 - this.b + (this.b * this.countWords(singleDocuemnt)) / this.averageLength)))
    );
  }

  search(queryWords: string[], topK?: number): BMDocument[] {
    queryWords = queryWords.filter((word) => word.length > 0);
    const wordCountMap = new Map<string, number>();
    queryWords.forEach((word) => {
      const count = wordCountMap.get(word) || 0;
      wordCountMap.set(word, count + 1);
    });
    const bmDocuments: BMDocument[] = this.documents.map((doc) => {
      const score = Array.from(wordCountMap).reduce(
        (acc, wordMap) =>
          acc +
          (this.score(wordMap[0], doc.documentText) * wordMap[1] * (1 + this.k3)) /
            (this.k3 + wordMap[1]),
        0
      );
      return { score, document: doc };
    });
    return bmDocuments.sort((a, b) => b.score - a.score).slice(0, topK || bmDocuments.length);
  }
}
