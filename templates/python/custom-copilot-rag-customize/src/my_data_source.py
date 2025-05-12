import os
from dataclasses import dataclass

from teams.ai.tokenizers import Tokenizer
from teams.ai.data_sources import DataSource
from teams.state.state import TurnContext
from teams.state.memory import Memory

@dataclass
class Result:
    output: str
    length: int
    too_long: bool

class MyDataSource(DataSource):
    """
    A data source that searches through a local directory of files for a given query.
    """

    def __init__(self, name):
        """
        Creates a new instance of the LocalDataSource instance.
        Initializes the data source.
        """
        self.name = name
        
        filePath = os.path.join(os.path.dirname(__file__), 'data')
        files = os.listdir(filePath)
        self._data = [open(os.path.join(filePath, file), 'r').read() for file in files]
        
    def name(self):
        return self.name

    async def render_data(self, context: TurnContext, memory: Memory, tokenizer: Tokenizer, maxTokens: int):
        """
        Renders the data source as a string of text.
        The returned output should be a string of text that will be injected into the prompt at render time.
        """
        query = memory.get('temp.input')
        if not query:
            return Result('', 0, False)
        
        result=''
        # Text search
        for data in self._data:
            if query in data:
                result += data
        # Key word search
        if 'history' in query.lower() or 'company' in query.lower():
            result += self._data[0]
        if 'perksplus' in query.lower() or 'program' in query.lower():
            result += self._data[1]
        if 'northwind' in query.lower() or 'health' in query.lower() or 'plan' in query.lower():
            result += self._data[2]
       
        return Result(self.formatDocument(result), len(result), False) if result!='' else Result('', 0, False)

    def formatDocument(self, result):
        """
        Formats the result string 
        """
        return f"<context>{result}</context>"