# Purpose
This add-in is used to test batching for custom functions with SharedRuntime. The add-in contains:
- Custom Functions
- Taskpane
- Uiless button to invoke 18000 CFs.

# How to use

1. Install  http-server: npm install -g http-server
2. Install  office-addin-https-reverse-proxy : npm install -g office-addin-https-reverse-proxy
3. Setup the Manifest.xml for devcatalog or FileShare
4. Run(Root of this folder): 'http-server -p  8080' to start the htpp server
5. Run: 'office-addin-https-reverse-proxy --url http://localhost:8080 --port 3000' to setup a https proxy
6. Open Excel workbook and Install the Add-in
7. Click on 'InvodeCF' or 'Run Action' button to run the 18000 Custom function all together.

# Maintainers
madhavagrawal17
