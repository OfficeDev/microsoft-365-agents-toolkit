// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import ts = require("typescript");
import { fetchRawFileContent } from "./utils";
import { SampleData } from "./samples/sampleData";
import { DocLinkTag, DocParagraph, DocPlainText, TSDocParser } from "@microsoft/tsdoc";

export class DeclarationFinder {
  private static DECLARATION_FILE_NAME = "office-js.d.ts";
  private static instance: DeclarationFinder;
  private definionFile: ts.SourceFile | undefined;
  private declarations: SampleData[] = [];

  private constructor() {}

  public static getInstance(): DeclarationFinder {
    if (!DeclarationFinder.instance) {
      DeclarationFinder.instance = new DeclarationFinder();
    }
    return DeclarationFinder.instance;
  }

  public async getClassSummariesForHost(host: string): Promise<SampleData[]> {
    await this.buildTypeDefAst();
    const sampleDatasOfHost: SampleData[] = this.declarations.filter((declaration) => {
      return (
        declaration.usage === host &&
        !!declaration.description &&
        !!declaration.definition &&
        !declaration.codeSample
      );
    });
    return sampleDatasOfHost;
  }

  public async getMethodsOrPropertiesForClass(
    host: string,
    className: string
  ): Promise<SampleData[]> {
    await this.buildTypeDefAst();
    const sampleDatasOfHost: SampleData[] = this.declarations.filter((declaration) => {
      return (
        declaration.usage === host &&
        declaration.definition === className &&
        !!declaration.description &&
        !!declaration.codeSample
      );
    });
    return sampleDatasOfHost;
  }

  private async buildTypeDefAst(): Promise<void> {
    if (!this.definionFile) {
      const typeDefStr = await fetchRawFileContent(
        `https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js/index.d.ts`
      );
      this.definionFile = ts.createSourceFile(
        DeclarationFinder.DECLARATION_FILE_NAME,
        typeDefStr,
        ts.ScriptTarget.Latest,
        true
      );

      this.buildDeclarationWithComments();
    }
  }

  private buildDeclarationWithComments() {
    // eslint-disable-next-line @typescript-eslint/no-this-alias
    const self = this;
    const sourceFile = this.definionFile;

    function visit(module: string | null, className: string | null, node: ts.Node) {
      if (ts.isModuleDeclaration(node)) {
        // The modules are Excel, Word, PPT, OfficeCore. etc.
        const moduleName = node.name?.getText();
        if (moduleName !== "Excel" && moduleName !== "Word" && moduleName !== "PowerPoint") {
          return;
        }
        ts.forEachChild(node, (node) => {
          visit(moduleName, className, node);
        });
      } else if (ts.isClassDeclaration(node)) {
        const clazzName = node.name?.getText() || null;
        const sampleData: SampleData = new SampleData(
          "", // name
          "", // docLink
          "", // codeSample
          "", // description
          clazzName ?? "", // definition
          module ?? "" // usage
        );
        const { summary } = self.getDocCommentAndSummary(node);
        sampleData.description = summary;
        self.declarations.push(sampleData);
        ts.forEachChild(node, (node) => {
          visit(module, clazzName, node);
        });
      } else if (
        ts.isInterfaceDeclaration(node) ||
        ts.isPropertyDeclaration(node) ||
        ts.isMethodDeclaration(node) ||
        ts.isEnumDeclaration(node)
      ) {
        const sampleData: SampleData = new SampleData(
          "", // name
          "", // docLink
          node.getText(), // codeSample
          "", // description
          className ?? "", // definition
          module ?? "" // usage
        );
        const { docComment, summary } = self.getDocCommentAndSummary(node);
        sampleData.description = summary;
        sampleData.docLink = docComment;
        self.declarations.push(sampleData);
      } else {
        ts.forEachChild(node, (node) => {
          visit(module, className, node);
        });
      }
    }

    ts.forEachChild(sourceFile!, (node) => {
      visit(null, null, node);
    });
  }

  private getDocCommentAndSummary(node: ts.Node): { docComment: string; summary: string } {
    // For the comments, we'd like to get the summary section of the comments. For example:
    // /**
    // * Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type.
    // * @param properties A JavaScript object with properties that are structured isomorphically to the properties of the object on which the method is called.
    // * @param options Provides an option to suppress errors if the properties object tries to set any read-only properties.
    // */
    // We expect to get the summary section of the comments, like:
    // "Sets multiple properties of an object at the same time. You can pass either a plain object with the appropriate properties, or another API object of the same type."
    const sourceFile = this.definionFile;
    const commentRanges = ts.getLeadingCommentRanges(sourceFile!.text, node.pos);
    const comments: string | undefined = commentRanges
      ? commentRanges
          .map((range) => sourceFile!.text.substring(range.pos, range.end).trim())
          .join("\n")
      : undefined;
    if (comments) {
      const tsDocParser = new TSDocParser();
      const tsDocComment = tsDocParser.parseString(comments).docComment;
      let description = "";
      const summarySectionIterator = tsDocComment?.summarySection.nodes.values();
      let summarySectionNext = summarySectionIterator.next();
      while (!summarySectionNext.done) {
        const node = summarySectionNext.value;
        if (node.kind === "PlainText") {
          // Deal with the plain text in the summary section. Like:
          // "Gets the first note item in this collection. Throws an `ItemNotFound` error if this collection is empty."
          description += (node as DocPlainText).text.trim().replace("`", "'") + " ";
        }
        if (node.kind === "LinkTag") {
          const link = node as DocLinkTag;
          if (link.linkText) {
            description += " " + link.linkText + " ";
          }
        }
        if (node.kind === "Paragraph") {
          // dealing with comments has extra content beyond pure text (link, for example), like:
          // "Contains a collection of {@link Word.NoteItem} objects."
          const paragraph = node as DocParagraph;
          const paragraphIterator = paragraph.nodes.values();
          let paragraphNext = paragraphIterator.next();
          while (!paragraphNext.done) {
            const paragraphNode = paragraphNext.value;
            if (paragraphNode.kind === "PlainText") {
              description +=
                (paragraphNode as unknown as DocPlainText).text.trim().replace("`", "'") + " ";
            }
            // dealing with links in the paragraph, like:
            // "{@link Word.NoteItem}"
            // It will get the Word.NoteItem from the link.
            if (paragraphNode.kind === "LinkTag") {
              const link = paragraphNode as DocLinkTag;
              let plainText = "";
              if (link.codeDestination) {
                const parts = link.codeDestination.memberReferences.map(
                  (memberReference) => memberReference.memberIdentifier?.identifier
                );
                plainText += parts.join(".");
              } else {
                plainText += link.linkText || "";
              }
              description += ` ${plainText} `;
            }
            paragraphNext = paragraphIterator.next();
          }
        }
        summarySectionNext = summarySectionIterator.next();
      }
      return { docComment: comments, summary: description };
    }
    return { docComment: "", summary: "" };
  }
}
