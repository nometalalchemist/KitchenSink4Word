# Glama inspection / generic container build.
# The live/COM tool family requires Windows + Microsoft Word and is not
# available in a container; the file-based majority of the toolset works
# anywhere. The server starts and serves MCP over stdio.
FROM python:3.12-slim
RUN pip install --no-cache-dir "kitchensink4word>=1.6.0"
ENTRYPOINT ["kitchensink4word"]
