// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { render } from "@inquirer/testing";
import figures from "figures";
import { select } from "../../../src/prompts/customizedListPrompt";
import { expect } from "../utils";

const choices = [
  { id: "id1", title: "title 1", detail: "detail 1" },
  { id: "id2", title: "title 2", detail: "detail 2" },
  { id: "id3", title: "title 3", detail: "detail 3" },
  { id: "id4", title: "title 4", detail: "detail 4" },
  { id: "id5", title: "title 5", detail: "detail 5" },
  { id: "id6", title: "title 6", detail: "detail 6" },
  { id: "id7", title: "title 7", detail: "detail 7" },
  { id: "id8", title: "title 8", detail: "detail 8" },
  { id: "id9", title: "title 9", detail: "detail 9" },
  { id: "id10", title: "title 10", detail: "detail 10" },
  { id: "id11", title: "title 11", detail: "detail 11" },
  { id: "id12", title: "title 12", detail: "detail 12" },
];

const getRadio = (checked: boolean) => {
  return checked ? figures.radioOn : figures.radioOff;
};

const trimOutput = (output: string) =>
  output
    .trim()
    .split("\n")
    .map((line) => line.trim().replace("( )", getRadio(false)).replace("(*)", getRadio(true)))
    .join("\n");

describe("select prompt", () => {
  it("use arrow keys to select an option", async () => {
    const { answer, events, getScreen } = await render(select, {
      message: "Select a string",
      choices,
    });

    expect(getScreen()).equal(
      trimOutput(`
        ? Select a string
        (*) title 1 detail 1
        ( ) title 2 detail 2
        ( ) title 3 detail 3
        ( ) title 4 detail 4
        ( ) title 5 detail 5
        ( ) title 6 detail 6
        ( ) title 7 detail 7
        (Use arrow keys to reveal more choices)`)
    );

    events.keypress("down");
    events.keypress("down");
    expect(getScreen()).equal(
      trimOutput(`
        ? Select a string
        ( ) title 1 detail 1
        ( ) title 2 detail 2
        (*) title 3 detail 3
        ( ) title 4 detail 4
        ( ) title 5 detail 5
        ( ) title 6 detail 6
        ( ) title 7 detail 7
        (Use arrow keys to reveal more choices)`)
    );

    events.keypress("enter");
    expect(getScreen()).equal("? Select a string title 3");

    expect(await answer).equal("id3");
  });

  it("allow setting a smaller page size", async () => {
    const { answer, events, getScreen } = await render(select, {
      message: "Select a string",
      choices,
      pageSize: 2,
    });

    expect(getScreen()).equal(
      trimOutput(`
        ? Select a string
        (*) title 1 detail 1
        ( ) title 2 detail 2
        (Use arrow keys to reveal more choices)`)
    );

    events.keypress("enter");
    expect(await answer).equal("id1");
  });

  it("allow setting a bigger page size", async () => {
    const { answer, events, getScreen } = await render(select, {
      message: "Select a string",
      choices,
      pageSize: 10,
    });

    expect(getScreen()).equal(
      trimOutput(`
        ? Select a string
        (*) title 1  detail 1
        ( ) title 2  detail 2
        ( ) title 3  detail 3
        ( ) title 4  detail 4
        ( ) title 5  detail 5
        ( ) title 6  detail 6
        ( ) title 7  detail 7
        ( ) title 8  detail 8
        ( ) title 9  detail 9
        ( ) title 10 detail 10
        (Use arrow keys to reveal more choices)`)
    );

    for (let i = 0; i < 8; i++) {
      events.keypress("down");
    }

    events.keypress("enter");
    expect(await answer).equal("id9");
  });

  it("allow setting the page size larger than choices length", async () => {
    const { answer, events, getScreen } = await render(select, {
      message: "Select a string",
      choices,
      pageSize: 14,
    });

    expect(getScreen()).equal(
      trimOutput(`
        ? Select a string
        (*) title 1  detail 1
        ( ) title 2  detail 2
        ( ) title 3  detail 3
        ( ) title 4  detail 4
        ( ) title 5  detail 5
        ( ) title 6  detail 6
        ( ) title 7  detail 7
        ( ) title 8  detail 8
        ( ) title 9  detail 9
        ( ) title 10 detail 10
        ( ) title 11 detail 11
        ( ) title 12 detail 12`)
    );

    for (let i = 0; i < 7; i++) {
      events.keypress("down");
    }

    events.keypress("enter");
    expect(await answer).equal("id8");
  });

  it("cycles through options", async () => {
    const { answer, events, getScreen } = await render(select, {
      message: "Select a string",
      choices,
      pageSize: 2,
      loop: true,
    });

    expect(getScreen()).equal(
      trimOutput(`
        ? Select a string
        (*) title 1 detail 1
        ( ) title 2 detail 2
        (Use arrow keys to reveal more choices)`)
    );

    events.keypress("up");
    events.keypress("up");

    expect(getScreen()).equal(
      trimOutput(`
        ? Select a string
        ( ) title 10 detail 10
        (*) title 11 detail 11
        (Use arrow keys to reveal more choices)`)
    );

    events.keypress("enter");
    expect(await answer).equal("id11");
  });
});
