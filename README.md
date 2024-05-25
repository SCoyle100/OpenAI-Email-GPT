# OpenAI Email GPT - An Intelligent Office Assistant

Welcome to my OpenAI Email GPT repository, an innovative project integrating the prowess of OpenAI's API with the versatility of Python. This project ambitiously bridges local functionalities and API-driven capabilities to streamline office tasks.

## Project Overview

At the heart of OpenAI Email GPT is a Python-based framework designed to interact seamlessly with both local system operations and external APIs. Key features include:

- **Local Function Integration**: Utilize Python's robustness to perform tasks such as dynamically creating folders within your system environment.

- **API Interactions**: Seamlessly integrates with powerful APIs like Microsoft Graph for Outlook operations and Pinecone for vector database management, enhancing the project's scope and efficiency.

## Vision and Goals

The driving vision behind OpenAI Email GPT is to develop a sophisticated office assistant, adept in handling both voice and text commands. This assistant is not just a passive tool but an active facilitator in office management. Core functionalities include:

- **Attachment Management**: Automated downloading and systematic organization of email attachments into designated folders.

- **Content Analysis**: Advanced analysis of attachments, leveraging AI to provide insights and summaries.

- **Script Execution**: Initiate and manage Python and PowerShell scripts via function calling, paving the way for contract automation and other office-related automations.

## Future Directions

As this project evolves, the goal is to refine these functionalities, ensuring a more intuitive, responsive, and efficient assistant. I envision a tool that not only simplifies but also anticipates the needs of the modern office, making workflows smoother and more integrated. I also plan on utilizing Streamlit for a user-friendly interface.

---

Your involvement and feedback are invaluable to this project. Feel free to explore, contribute, and be a part of crafting the future of office automation with OpenAI Email GPT.

Updates - 2/1/2024:

I am planning on doing the following in the near future:

- Modularizing script and reducing technical debt
- Adding more functions such as triggering python, powershell or UiPath automations
- Utilizing GPT Vision
- Creating a separate repository that is essentially the same project, but instead uses the OpenAI Assistants API (which is currently a bit too expensive for production use).  This would ultimately allow for email attachments like Excel files to be downloaded and analyzed with code interpreter, and then decisions would be be made based on information from a vector database (either Pinecone or ChromaDB).  


Updates - 5/17/2024:

- I have successfully refactored from a procedural to an object-oriented project.  This will help with scalability, maintainability, unit-testing, and more.  The procedural project is saved in branch "archive_procedural".


-----------------------------------------------------------------




