from fastmcp import FastMCP
import random
import string

mcp = FastMCP("ttk")

@mcp.tool
def createRandomName(pre: str = "fxui"):
    random_chars = ''.join(random.choices(string.ascii_lowercase + string.digits, k=5))
    return pre + random_chars

if __name__ == "__main__":
    mcp.run()