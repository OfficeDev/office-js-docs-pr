:: Setup
:: -----

setlocal enabledelayedexpansion

IF NOT DEFINED SITE (
  SET SITE=%~dp0%..
)

IF NOT DEFINED DEPLOYMENT_SOURCE (
  SET DEPLOYMENT_SOURCE=%SITE%\repository
)

IF NOT DEFINED DEPLOYMENT_TARGET_DIR (
  SET DEPLOYMENT_TARGET_DIR=%SITE%\wwwroot\OfficeDocuments
)

IF NOT DEFINED DEPLOYMENT_TEMPLATE (
  SET DEPLOYMENT_TEMPLATE=%SITE%\wwwroot\MDConverter\html-template
)

IF NOT DEFINED APIDOCS_PATH (
  SET APIDOCS_PATH=%SITE%\wwwroot\MDConverter\bin
)

%APIDOCS_PATH%\apidocs.exe publish --path %DEPLOYMENT_SOURCE%\docs --output %DEPLOYMENT_TARGET_DIR%\docs --template %DEPLOYMENT_TEMPLATE% --format mustache --insert-gitInfo true --gitUrl https://github.com/OfficeDev/office-js-docs/tree/master

%APIDOCS_PATH%\apidocs.exe publish --path %DEPLOYMENT_SOURCE%\reference --output %DEPLOYMENT_TARGET_DIR%\reference --template %DEPLOYMENT_TEMPLATE% --format mustache --insert-gitInfo true --gitUrl https://github.com/OfficeDev/office-js-docs/tree/master

xcopy %DEPLOYMENT_SOURCE%\images %DEPLOYMENT_TARGET_DIR%\images -recurse



