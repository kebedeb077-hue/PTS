**Preparing and Deploying the Publications Tracking System Website**

- **Source folder**: `Publications_Tracking_System`
- **Packaging script**: `scripts/package_website.sh` — creates `deploy/Publications_Tracking_System.zip` and writes `deploy/mdb_files.txt` listing any `.mdb` files.

Checklist before publishing:
- **Sensitive DB files**: The site contains `.mdb` files (Access). Review `deploy/mdb_files.txt` and ensure the deployment target has appropriate protections. Consider moving databases to a protected path outside web root and updating connection strings in ASP files.
- **IIS / ASP Classic**: This site is classic ASP (`.asp` files). Deploy to a Windows Server with IIS and enable Classic ASP. Ensure correct application pool settings and that the account used by IIS has file read/write access where required.
- **Permissions**: Give the webserver account read access to site files; give write access only where the application expects to write (e.g. upload or log directories). Do NOT make the entire site writable.
- **Remove dev artifacts**: The packaging script excludes `_vti_cnf` and `_vti_pvt` folders and common extras. Remove any other local development files not intended for production.
- **Backups**: Backup `Forms.mdb` and `FormsEm.mdb` before moving them; keep a copy offline.
- **Testing**: After deployment, test key pages (login/change password, forms, reports) and the database connections.

Quick packaging and inspection commands:

```bash
# create the zip artifact
scripts/package_website.sh

# inspect contents
unzip -l deploy/Publications_Tracking_System.zip | less

# view the found .mdb files
cat deploy/mdb_files.txt
```

If you want, I can also add a CI job to create this artifact automatically (GitHub Actions), or create a cleaned copy of the site with databases removed. Tell me which you prefer.
