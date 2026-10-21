"""m365graph: a small Microsoft Graph client, and SharePoint libraries as stores of folders and
files. It knows nothing of the projects that use it. Contract: docs/public_contract.md."""

from m365graph.client import GraphClient, GraphError
from m365graph.drives import DriveStore, Folder, PathError, TooLarge, find_site, libraries

__version__ = "0.1.0"

__all__ = ["DriveStore", "Folder", "GraphClient", "GraphError", "PathError", "TooLarge",
           "__version__", "find_site", "libraries"]
