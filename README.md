# Azure Data Factory Pipeline Run Exporter

## Overview

This application allows you to retrieve Azure Data Factory pipeline run details using the Azure Resource Manager API and export them to an Excel file for analysis and reporting.

## Features

- Authentication using Azure AD application client secret.
- Retrieval of pipeline run details using Azure Resource Manager API.
- Export of pipeline run details to an Excel file.

## Prerequisites

- Azure subscription with an Azure Data Factory instance.
- Azure AD application with client secret authentication method.

## Setup

1. Clone the repository or download the source code.
2. Open the solution in Visual Studio.
3. Install the necessary NuGet packages (`Microsoft.Azure.Management.DataFactory`, `Microsoft.Identity.Client`, `Newtonsoft.Json`, `EPPlus`).
4. Replace the `tenantId`, `clientId`, `clientSecret`, and `url` variables with your Azure AD and Azure Data Factory details.

## Usage

1. Run the application.
2. The application will authenticate using the provided Azure AD application credentials.
3. It will then retrieve the pipeline run details from the specified Azure Data Factory instance.
4. The details will be exported to an Excel file named `run.xlsx` in the project directory.

## Code Structure

- `Program.cs`: Contains the `Main` method and entry point for the application.
- `DataFactoryPipelineRunInfo.cs`: Defines the model classes for pipeline run details.
- `Root.cs`: Defines the root model class for deserializing API responses.
- `ExportDataToExcel` method: Exports the pipeline run details to an Excel file.
- `CollectRunDetails` method: Collects pipeline run details from the Azure Data Factory API.

## Limitations

- The application currently retrieves all pipeline run details within a specified date range. It does not support filtering by specific pipeline names or statuses.

## Future Enhancements

- Add support for filtering pipeline run details by pipeline name or status.
- Improve error handling and logging.
- Support for exporting to different file formats (e.g., CSV, JSON).

