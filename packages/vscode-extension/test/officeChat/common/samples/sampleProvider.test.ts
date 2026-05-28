import { expect } from "chai";
import { SampleProvider } from "../../../../src/officeChat/common/samples/sampleProvider";
import * as utils from "../../../../src/chat/utils";
import sinon from "ts-sinon";
import { Spec } from "../../../../src/officeChat/common/skills/spec";

describe("SampleProvider", () => {
  const sandbox = sinon.createSandbox();

  beforeEach(() => {});

  afterEach(() => {
    sandbox.restore();
  });

  it("top K most relevant scenario sample codes LLM", async () => {
    sandbox
      .stub(utils, "getCopilotResponseAsString")
      .resolves('{"picked":["description1", "description2"]}');
    const sample = "a fake code sample";
    const scenario = "insert annotation into document";
    const host = "Word";
    const spec = new Spec("some user input");
    const topKSamples = await SampleProvider.getInstance().getMostRelevantDeclarationsUsingLLM(
      null as any,
      host,
      scenario,
      sample,
      spec
    );

    expect(topKSamples).to.exist;
    expect(topKSamples).to.be.an("map");
    // Add more assertions based on what you expect the topKSamples to be
  });

  it("top K most relevant scenario sample codes BM25", async () => {
    const k = 2;
    const scenario = "insert annotation into document";
    const host = "Word";
    let topKSamples = await SampleProvider.getInstance().getTopKMostRelevantScenarioSampleCodesBM25(
      null as any,
      host,
      scenario,
      k
    );
    if (topKSamples.size === 0) {
      topKSamples = await SampleProvider.getInstance().getTopKMostRelevantScenarioSampleCodesBM25(
        null as any,
        host,
        scenario,
        k
      );
    }
    expect(topKSamples).to.exist;
    expect(topKSamples).to.be.an("map");
    expect(topKSamples).to.have.lengthOf(k);
    // Add more assertions based on what you expect the topKSamples to be
  });

  it("not valid host BM25", async () => {
    const k = 2;
    const scenario = "insert annotation into document";
    const host = "FakeHost";
    const topKSamples =
      await SampleProvider.getInstance().getTopKMostRelevantScenarioSampleCodesBM25(
        null as any,
        host,
        scenario,
        k
      );
    expect(topKSamples).to.exist;
    expect(topKSamples).to.be.an("map");
    expect(topKSamples).to.have.lengthOf(0);
    // Add more assertions based on what you expect the topKSamples to be
  });

  it("no class available for given host sample codes LLM", async () => {
    const sample = "a fake code sample";
    const scenario = "insert annotation into document";
    const host = "UnkownHost";
    const spec = new Spec("some user input");
    const topKSamples = await SampleProvider.getInstance().getMostRelevantDeclarationsUsingLLM(
      null as any,
      host,
      scenario,
      sample,
      spec
    );

    expect(topKSamples).to.exist;
    expect(topKSamples).to.be.an("map");
    expect(topKSamples).to.have.lengthOf(0);
    // Add more assertions based on what you expect the topKSamples to be
  });

  it("token overload relevant to scenario sample codes LLM", async () => {
    const sample = "a fake code sample";
    const scenario = `
    Video provides a powerful way to help you prove your point. When you click Online Video, you can paste in the embed code for the video you want to add. You can also type a keyword to search online for the video that best fits your document.
To make your document look professionally produced, Word provides header, footer, cover page, and text box designs that complement each other. For example, you can add a matching cover page, header, and sidebar. Click Insert and then choose the elements you want from the different galleries.
Themes and styles also help keep your document coordinated. When you click Design and choose a new Theme, the pictures, charts, and SmartArt graphics change to match your new theme. When you apply styles, your headings change to match the new theme.
Save time in Word with new buttons that show up where you need them. To change the way a picture fits in your document, click it and a button for layout options appears next to it. When you work on a table, click where you want to add a row or a column, and then click the plus sign.
Reading is easier, too, in the new Reading view. You can collapse parts of the document and focus on the text you want. If you need to stop reading before you reach the end, Word remembers where you left off - even on another device.
    `;
    const host = "UnkownHost";
    const spec = new Spec("some user input");
    const topKSamples = await SampleProvider.getInstance().getMostRelevantDeclarationsUsingLLM(
      null as any,
      host,
      scenario.repeat(100), // repeat the scenario to make it longer
      sample,
      spec
    );

    expect(topKSamples).to.exist;
    expect(topKSamples).to.be.an("map");
    expect(topKSamples).to.have.lengthOf(0);
    // Add more assertions based on what you expect the topKSamples to be
  });

  it("no class relevant to scenario sample codes LLM", async () => {
    sandbox.stub(utils, "getCopilotResponseAsString").resolves('{"picked":[]}');
    const sample = "a fake code sample";
    const scenario = "insert annotation into document";
    const host = "UnkownHost";
    const spec = new Spec("some user input");
    const topKSamples = await SampleProvider.getInstance().getMostRelevantDeclarationsUsingLLM(
      null as any,
      host,
      scenario,
      sample,
      spec
    );

    expect(topKSamples).to.exist;
    expect(topKSamples).to.be.an("map");
    expect(topKSamples).to.have.lengthOf(0);
    // Add more assertions based on what you expect the topKSamples to be
  });

  it("token overload", async () => {
    sandbox.stub(utils, "getCopilotResponseAsString").resolves('{"picked":[]}');
    const sample = "a fake code sample";
    const scenario = "insert annotation into document";
    const host = "Excel";
    const spec = new Spec("some user input");
    sandbox
      .stub(utils, "countMessagesTokens")
      .onFirstCall()
      .returns(4000)
      .onSecondCall()
      .returns(4000);
    const topKSamples = await SampleProvider.getInstance().getMostRelevantDeclarationsUsingLLM(
      null as any,
      host,
      scenario,
      sample,
      spec
    );

    expect(topKSamples).to.exist;
    expect(topKSamples).to.be.an("map");
    expect(topKSamples).to.have.lengthOf(0);
    // Add more assertions based on what you expect the topKSamples to be
  });

  it("no methods or properties relevant to scenario sample codes LLM", async () => {
    const getCopilotResponseAsStringStub = sandbox.stub(utils, "getCopilotResponseAsString");
    getCopilotResponseAsStringStub
      .onCall(0)
      .returns(Promise.resolve('{"picked":["Workbook", "Worksheet", "Range", "Chart", "Shape"]}'));
    getCopilotResponseAsStringStub.onCall(1).returns(Promise.resolve('{"picked":[]}'));
    const sample = "a fake code sample";
    const scenario = "insert annotation into document";
    const host = "Excel";
    const spec = new Spec("some user input");
    const topKSamples = await SampleProvider.getInstance().getMostRelevantDeclarationsUsingLLM(
      null as any,
      host,
      scenario,
      sample,
      spec
    );

    expect(topKSamples).to.exist;
    expect(topKSamples).to.be.an("map");
    expect(topKSamples).to.have.lengthOf(0);
    // Add more assertions based on what you expect the topKSamples to be
  });

  it("no class returned from LLM", async () => {
    const getCopilotResponseAsStringStub = sandbox.stub(utils, "getCopilotResponseAsString");
    getCopilotResponseAsStringStub.onCall(0).returns(Promise.resolve('{"picked":[]}'));
    const sample = "a fake code sample";
    const scenario = "insert annotation into document";
    const host = "Excel";
    const spec = new Spec("some user input");
    const topKSamples = await SampleProvider.getInstance().getMostRelevantDeclarationsUsingLLM(
      null as any,
      host,
      scenario,
      sample,
      spec
    );

    expect(topKSamples).to.exist;
    expect(topKSamples).to.be.an("map");
    expect(topKSamples).to.have.lengthOf(0);
    // Add more assertions based on what you expect the topKSamples to be
  });

  it("one methods or properties relevant to scenario sample codes LLM", async () => {
    const getCopilotResponseAsStringStub = sandbox.stub(utils, "getCopilotResponseAsString");
    getCopilotResponseAsStringStub.onCall(0).returns(Promise.resolve('{"picked":["Shape"]}'));
    getCopilotResponseAsStringStub
      .onCall(1)
      .returns(
        Promise.resolve('{"picked":["class: Shape; readonly connectionSiteCount: number;"]}')
      );
    const sample = "a fake code sample";
    const scenario = "insert annotation into document";
    const host = "Excel";
    const spec = new Spec("some user input");
    const topKSamples = await SampleProvider.getInstance().getMostRelevantDeclarationsUsingLLM(
      null as any,
      host,
      scenario,
      sample,
      spec
    );

    expect(topKSamples).to.exist;
    expect(topKSamples).to.be.an("map");
    expect(topKSamples).to.have.lengthOf(1);
    // Add more assertions based on what you expect the topKSamples to be
  });

  it("mutiple methods or properties relevant to scenario sample codes LLM", async () => {
    const getCopilotResponseAsStringStub = sandbox.stub(utils, "getCopilotResponseAsString");
    getCopilotResponseAsStringStub.onCall(0).returns(Promise.resolve('{"picked":["Shape"]}'));
    getCopilotResponseAsStringStub
      .onCall(1)
      .returns(
        Promise.resolve(
          '{"picked":["class: Shape; readonly connectionSiteCount: number; altTextDescription: string;"]}'
        )
      );
    const sample = "a fake code sample";
    const scenario = "insert annotation into document";
    const host = "Excel";
    const spec = new Spec("some user input");
    const topKSamples = await SampleProvider.getInstance().getMostRelevantDeclarationsUsingLLM(
      null as any,
      host,
      scenario,
      sample,
      spec
    );

    expect(topKSamples).to.exist;
    expect(topKSamples).to.be.an("map");
    // Add more assertions based on what you expect the topKSamples to be
  });

  it("return method without class to scenario sample codes LLM", async () => {
    const getCopilotResponseAsStringStub = sandbox.stub(utils, "getCopilotResponseAsString");
    getCopilotResponseAsStringStub.onCall(0).returns(Promise.resolve('{"picked":["Shape"]}'));
    getCopilotResponseAsStringStub
      .onCall(1)
      .returns(Promise.resolve('{"picked":["readonly connectionSiteCount: number;"]}'));
    const sample = "a fake code sample";
    const scenario = "insert annotation into document";
    const host = "Excel";
    const spec = new Spec("some user input");
    const topKSamples = await SampleProvider.getInstance().getMostRelevantDeclarationsUsingLLM(
      null as any,
      host,
      scenario,
      sample,
      spec
    );

    expect(topKSamples).to.exist;
    expect(topKSamples).to.be.an("map");
    expect(topKSamples).to.have.lengthOf(1);
    // Add more assertions based on what you expect the topKSamples to be
  });

  it("giant class to scenario sample codes LLM", async () => {
    const getCopilotResponseAsStringStub = sandbox.stub(utils, "getCopilotResponseAsString");
    getCopilotResponseAsStringStub
      .onCall(0)
      .returns(
        Promise.resolve(
          '{"picked":["Workbook", "Worksheet", "Range", "FunctionResult", "Functions"]}'
        )
      );
    getCopilotResponseAsStringStub
      .onCall(1)
      .returns(
        Promise.resolve(
          '{"picked":["class: Workbook; readonly application: Excel.Application;", "class: Worksheet; readonly charts: Excel.ChartCollection;"]}'
        )
      );
    getCopilotResponseAsStringStub
      .onCall(2)
      .returns(
        Promise.resolve(
          '{"picked":["class: Functions; arabic(text: string | Excel.Range | Excel.RangeReference | Excel.FunctionResult<any>): FunctionResult<number>;"]}'
        )
      );
    getCopilotResponseAsStringStub
      .onCall(3)
      .returns(Promise.resolve('{"picked":["class: 3; method 1;"]}'));
    getCopilotResponseAsStringStub
      .onCall(4)
      .returns(Promise.resolve('{"picked":["class: 3; method 2;"]}'));
    getCopilotResponseAsStringStub
      .onCall(5)
      .returns(Promise.resolve('{"picked":["class: 3; method 3;"]}'));
    getCopilotResponseAsStringStub
      .onCall(6)
      .returns(Promise.resolve('{"picked":["class: 3; method 4;"]}'));
    getCopilotResponseAsStringStub
      .onCall(7)
      .returns(Promise.resolve('{"picked":["class: 3; method 5;"]}'));
    getCopilotResponseAsStringStub
      .onCall(8)
      .returns(Promise.resolve('{"picked":["class: 3; method 6;"]}'));
    getCopilotResponseAsStringStub
      .onCall(9)
      .returns(Promise.resolve('{"picked":["class: 3; method 7;"]}'));

    const sample = "";
    const scenario =
      "To set up streaming custom functions with the Office JS API that fetch real-time data from the web at 10-second intervals, you should follow these steps: 1. Define a function in a JavaScript or Typescript file that fetches the data from the web. 2. Ensure this function is async and is continuously running with a call every 10 seconds. 3. In the custom functions metadata, register this function as a streaming function. 4. Test this function in Excel to confirm it behaves correctly.";
    const host = "Excel";
    const spec = new Spec("some user input");
    const topKSamples = await SampleProvider.getInstance().getMostRelevantDeclarationsUsingLLM(
      null as any,
      host,
      scenario,
      sample,
      spec
    );

    expect(topKSamples).to.exist;
    expect(topKSamples).to.be.an("map");
  }).timeout(10000);
});
