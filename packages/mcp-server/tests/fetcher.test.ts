import { expect } from 'chai';
import * as sinon from 'sinon';
import nock from 'nock';
import { fetchSchema, SchemaType } from '../src/fetcher';

describe('Fetcher', () => {
  // Setup and teardown
  beforeEach(() => {
    // Enable nock for HTTP request mocking
    nock.disableNetConnect();
  });

  afterEach(() => {
    // Clean up nock after each test
    nock.cleanAll();
    nock.enableNetConnect();
    sinon.restore();
  });
  describe('fetchSchema', () => {
    it('should fetch app_manifest schema successfully', async () => {
      const schemaVersion = 'v1.16';
      const mockSchemaContent = {
        $schema: 'http://json-schema.org/draft-07/schema#',
        type: 'object',
        properties: {
          // Mock properties
        }
      };
      
      // Mock the HTTP request
      nock('https://developer.microsoft.com')
        .get(`/json-schemas/teams/${schemaVersion}/MicrosoftTeams.schema.json`)
        .reply(200, mockSchemaContent);
      
      const result = await fetchSchema('app_manifest' as SchemaType, schemaVersion);
      const parsed = JSON.parse(result);
      
      expect(parsed).to.have.property('schema_url');
      expect(parsed.schema_url).to.equal(`https://developer.microsoft.com/json-schemas/teams/${schemaVersion}/MicrosoftTeams.schema.json`);
      expect(parsed).to.have.property('content');
      expect(parsed.content).to.deep.equal(mockSchemaContent);
    });

    it('should fetch declarative_agent_manifest schema successfully', async () => {
      const schemaVersion = 'v1.0';
      const mockSchemaContent = {
        $schema: 'http://json-schema.org/draft-07/schema#',
        type: 'object',
        properties: {
          // Mock properties
        }
      };
      
      // Mock the HTTP request
      nock('https://developer.microsoft.com')
        .get(`/json-schemas/copilot/declarative-agent/${schemaVersion}/schema.json`)
        .reply(200, mockSchemaContent);
      
      const result = await fetchSchema('declarative_agent_manifest' as SchemaType, schemaVersion);
      const parsed = JSON.parse(result);
      
      expect(parsed).to.have.property('schema_url');
      expect(parsed.schema_url).to.equal(`https://developer.microsoft.com/json-schemas/copilot/declarative-agent/${schemaVersion}/schema.json`);
      expect(parsed).to.have.property('content');
      expect(parsed.content).to.deep.equal(mockSchemaContent);
    });

    it('should fetch api_plugin_manifest schema successfully', async () => {
      const schemaVersion = 'v1.0';
      const mockSchemaContent = {
        $schema: 'http://json-schema.org/draft-07/schema#',
        type: 'object',
        properties: {
          // Mock properties
        }
      };
      
      // Mock the HTTP request
      nock('https://developer.microsoft.com')
        .get(`/json-schemas/copilot/plugin/${schemaVersion}/schema.json`)
        .reply(200, mockSchemaContent);
      
      const result = await fetchSchema('api_plugin_manifest' as SchemaType, schemaVersion);
      const parsed = JSON.parse(result);
      
      expect(parsed).to.have.property('schema_url');
      expect(parsed.schema_url).to.equal(`https://developer.microsoft.com/json-schemas/copilot/plugin/${schemaVersion}/schema.json`);
      expect(parsed).to.have.property('content');
      expect(parsed.content).to.deep.equal(mockSchemaContent);
    });

    it('should return cached schema if already fetched', async () => {
      const schemaVersion = 'v1.16';
      const mockSchemaContent = {
        $schema: 'http://json-schema.org/draft-07/schema#',
        type: 'object',
        properties: {
          // Mock properties
        }
      };
      
      // Mock the HTTP request only once - it should not be called for the second fetch
      nock('https://developer.microsoft.com')
        .get(`/json-schemas/teams/${schemaVersion}/MicrosoftTeams.schema.json`)
        .reply(200, mockSchemaContent);
        // First fetch
      await fetchSchema('app_manifest' as SchemaType, schemaVersion);
      
      // Second fetch - should use cache
      const result = await fetchSchema('app_manifest' as SchemaType, schemaVersion);
      const parsed = JSON.parse(result);
      
      expect(parsed).to.have.property('schema_url');
      expect(parsed).to.have.property('content');
      expect(parsed.content).to.deep.equal(mockSchemaContent);
    });

    it('should return error message when schema fetch fails', async () => {
      const schemaVersion = 'non-existent';
      
      // Mock the HTTP request with a 404 response
      nock('https://developer.microsoft.com')
        .get(`/json-schemas/teams/${schemaVersion}/MicrosoftTeams.schema.json`)
        .reply(404);
        const result = await fetchSchema('app_manifest' as SchemaType, schemaVersion);
      
      expect(result).to.include('Failed fetching schema');
      expect(result).to.include('HTTP error with status: 404');
    });
  });
});
