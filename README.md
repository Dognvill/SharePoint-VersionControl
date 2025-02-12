## SharePoint Version Management and Cleanup Script
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

### Features
- Enable Intelligent Versioning at tenant level for new SharePoint sites
- Enable Intelligent Versioning for existing SharePoint sites
- Clean up version history based on age (days)
- Clean up version history based on version count
- Interactive site selection (single, multiple, or all sites)

### Prerequisites
- PowerShell 5.1 or higher
- SharePoint Online Management Shell (automatically installed if missing)
- SharePoint Administrator permissions

### License
This script is released under the MIT License. See the LICENSE file for details.
