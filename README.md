## SharePoint Version Management
A PowerShell automation solution for implementing Intelligent Versioning and managing storage effectively in SharePoint Online environments. This script helps organizations optimize their SharePoint storage usage while maintaining essential version history.

### The Challenge
SharePoint Online's versioning feature, while essential for collaboration and document management, can lead to significant storage consumption over time. By default, SharePoint retains all document versions indefinitely, which can result in:
- Excessive storage usage and increased costs
- Reduced system performance
- Unnecessary retention of outdated content

### Solution: Intelligent Versioning
Intelligent Versioning is Microsoft's strategic approach to version management that automatically optimizes storage while preserving important document history. The system implements "automatic version thinning" based on the following timeline:
- First 30 Days: Preserves all versions
- 30-180 Days: Maintains daily versions
- Beyond 180 Days: Keeps weekly versions

For organizations requiring more granular control, manual version management can be implemented through PowerShell to set specific version limits per document library, remove versions older than a defined date, and selectively trim minor versions while preserving major versions. 

## Retention Policies and Preservation
### Retention Policies
SharePoint Online uses retention policies to manage content lifecycle:
- In-Place Hold: Retains content without user visibility
- Litigation Hold: Preserves content for legal requirements
- Retention Labels: Applies specific retention rules to individual items
- Site Policies: Manages retention at the site collection level

### Preservation Hold Library
The Preservation Hold Library is a hidden document library created when retention policies are applied:
- Automatically captures deleted content and version history
- Maintains copies of modified content
- Stores preserved content in its original format
- Retains all metadata and audit trails

This script provides tools to:
- Review Preservation Hold Library contents
- Download preserved content with metadata
- Export to Azure Blob Storage with full fidelity via SAS
- Generate detailed inventory reports

### Features
- Enable Intelligent Versioning at tenant level for new SharePoint sites
- Enable Intelligent Versioning for existing SharePoint sites
- Clean up version history based on age (days)
- Clean up version history based on version count
- Interactive site selection (single, multiple, or all sites)
- Preservation Hold Library management:
    - Content inventory and reporting
    - Bulk content download
    - Azure Blob Storage integration
    - Metadata preservation

## Authentication Options
The script employs different authentication methods depending on the operation being performed.
#### SharePoint Online Management Shell (SPO Services)
Core SharePoint management operations (site selection and automatic versioning) use SPO Services with interactive login. This requires SharePoint Admin permissions and handles all tenant-level configurations.

#### PnP PowerShell Authentication
Preservation Hold Library operations require PnP PowerShell authentication, which supports two methods:
- Delegated Authentication (Interactive): Uses Azure AD App registration with Client ID with delegated permissions. Uses modern authentication flow with user interaction, combining app permissions with user's rights
- Application Authentication (App-only): Uses Azure AD App registration with Client ID and Certificate, defining permissions at the application level for automated processes

_Note: For information about PnP PowerShell authentication methods, refer to the PnP Documentation: https://pnp.github.io/powershell/articles/authentication.html_

#### Azure Blob Storage Access
When using the optional cloud storage integration feature, the script requires SAS (Shared Access Signature) token authentication with Blob Core permissions from your Storage Account.

### Prerequisites
- PowerShell 7
- SharePoint Online Management Shell (automatically installed if missing)
- SharePoint Administrator permissions

### License
This script is released under the MIT License. See the LICENSE file for details.
