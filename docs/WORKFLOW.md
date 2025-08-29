# Workflow – VBA Automation Suite

## Objective
Define the **end-to-end workflow** for maintaining, testing, and deploying VBA automation modules in this repository.

---

## Development Workflow

1. **Design** → Identify task to automate (Outlook, Excel, ISO).
2. **Build** → Write VBA macro in a development `.xlsm` file.
3. **Test** → Run against sample data (attachments, reports, dummy templates).
4. **Commit** → Push updates to `dev` branch.
5. **Review** → Peer or self-review for compliance and performance.
6. **Merge** → Approved code merged into `main`.

---

## Deployment Workflow

1. Use **Push-updates-to-templates-ISO** for distributing code to all templates.  
2. Run **PushProject2Production** to finalize build and move to Production folder.  
3. Lock VB projects for ISO compliance before release.  
4. Document changes in `CHANGELOG.md`.

---

## User Workflow

- **Analysts** → Run data macros (`hide-unhide-delete-hidden-column`, `Loopfiles-Analyse-Print`).  
- **Operations** → Run `Download_attachments` daily/weekly to collect files.  
- **ISO/QA Teams** → Run bulk signature, lock/unlock, and production push.  
- **Developers** → Handle Gemini API testing and template updates.

---

## Maintenance Workflow

- Backup all `.xlsm` templates weekly.  
- Rotate API keys periodically.  
- Clean `Production` folder every quarter to remove deprecated builds.  
- Review code comments to ensure ISO wording is up-to-date.  

---

## Future Expansion

- API integration with multiple AI providers.  
- CI/CD integration for VBA deployments.  
- Centralized logging system for all macros.  

