# IFilterExtractor
A simple component to extract just the text from any file that has an IFilter installed. Available as a C++ COM component and as a C# .NET library.

IFilter is a COM component and .Net test jig that uses installed IFilter providers to extract the text from any file. The providers for the various formats are available from most vendors as well as a couple third-party providers. The IFilter providers are used by Microsoft Index Server, Microsoft Sharepoint Server and Microsoft Desktop Search to extract the indexable text for a file. By using the same interfaces, it is possible to extract just the text (less formatting) from just about any file from Microsoft Word .DOC files to .MP3 files.
