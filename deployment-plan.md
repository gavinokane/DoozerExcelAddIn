# Deployment Plan for DoozerExcelAddIn Web App

## Progress So Far
1. **Project Initialization**:
   - Initialized a local Git repository.
   - Added all project files to the repository.
   - Committed the initial changes with the message "Initial commit."

2. **GitHub CLI Installation**:
   - Installed GitHub CLI using the `winget` package manager.
   - Updated the system PATH to include the GitHub CLI directory.

3. **Issues Encountered**:
   - Despite adding GitHub CLI to the PATH, it is still not recognized in the terminal.
   - Attempts to authenticate using `gh auth login` were unsuccessful.

## Next Steps
1. **Verify GitHub CLI Installation**:
   - Restart VSCode and ensure the PATH update is effective.
   - Retry the `gh auth login` command to authenticate with GitHub.

2. **Create GitHub Repository**:
   - Use GitHub CLI or manual methods to create a new repository named `DoozerExcelAddIn`.
   - Push the local repository to GitHub.

3. **Configure Azure Deployment**:
   - Set up GitHub as the deployment source for the Azure App Service.
   - Use the following command:
     ```
     az webapp deployment source config --name doozeraiexceladdin --resource-group rg-prod01 --repo-url https://github.com/gavinokane/DoozerExcelAddIn --branch main --manual-integration
     ```

4. **Test Deployment**:
   - Verify that the application is successfully deployed and accessible via the Azure App Service URL.
   - Check logs for any issues and resolve them as needed.

5. **Finalize Deployment**:
   - Configure environment variables and application settings in Azure.
   - Enable monitoring and logging for the web app.

## Notes
- If GitHub CLI continues to fail, consider using manual methods to create the repository and authenticate with GitHub.
- Ensure all dependencies are installed and the deployment package is correctly structured.