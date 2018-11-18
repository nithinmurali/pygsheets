
Changelog
=========


Version 2.0.0
-------------
This version is not backwards compatible with 1.x
There is major rework in the library with this release.
Some functions are renamed to have better consistency in naming and clear meaning.

- update_cell() renamed to update_value()
- update_cells() renamed to update_values()
- update_cells_prop() renamed to update_cells()
- changed authorize() params : outh_file -> client_secret, outh_creds_store ->credentials_directory, service_file -> service_account_file, credentials -> custom_credentials
- teamDriveId, enableTeamDriveSupport changed to client.drive.enable_team_drive, include_team_drive_items
- parameter changes for all get_* functions : include_empty, include_all changed to include_tailing_empty, include_tailing_empty_rows
- parameter changes in created_protected_range() : gridrange param changed to start, end
- remoed batch mode
- find() splited into find() and replace()
- removed (show/hide)_(row/column), use (show/hide)_dimensions instead
- removed link/unlink from spreadsheet

**New Features added**
- chart Support added
- Sort feature added
- Better support for protected ranges
- Multi header/index support in dataframes
- Removed the dependency on oauth2client and uses google-auth and google-auth-oauth.

Other bug fixes and performance improvements


Version 1.1.4
-------------
Changelog not available