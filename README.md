# Document Listing Web Part

## Summary

This SharePoint Framework (SPFx) web part provides a rich, categorized document listing experience. It is designed to allow users to navigate documents using a two-level hierarchy (Category and Sub-Category), search for documents, and request access to specific files. The solution integrates with a secondary list to track access requests and download counts.

![Solution Mockup](assets/document_listing_mockup.png)

## Features

- **Categorized Navigation**: 
  - **Left Navigation**: Filter documents by a primary "Category" column.
  - **Top Tabs**: Further filter documents within a category using a "Sub-Category" column.
- **Sorting**: Sort documents by Title, Description, or Modified Date (defaulting to Modified Date Descending).
- **Pagination**: Configurable pagination to manage large sets of documents (default 10 items per page).
- **Request Access Workflow**:
  - Users can click a "Request Access" (mail icon) button on any document.
  - Requests are logged to a configured SharePoint List.
  - Supports tracking request counts (incrementing a counter if the user requests the same file multiple times).
  - Toast notifications provide immediate feedback on request status.
- **Theme Awareness**: The web part automatically adapts to the current SharePoint site theme.

## Configuration

The web part is fully configurable via the Property Pane.

![Configuration Pane Mockup](assets/webpart_configuration_mockup.png)

### 1. General Configuration
- **Web Part Title**: The header text displayed at the top of the web part.
- **Document Library Name**: Select the source Document Library (must be a library, template 101).
- **Category Column**: Select the Choice column to use for the left-hand navigation.
- **Sub Category Column**: Select the Choice column to use for the top tab navigation.
- **Title Field**: (Optional) specific text column to display as the document title.
- **Description Field**: (Optional) specific text column to display as the document description.
- **Max Rows per Page**: Slider to set the number of items per page (Range: 5-100, Default: 10).

### 2. Request Access Configuration
To enable the "Request Access" feature, you must have a separate generic SharePoint List (template 100) created to store the requests.

- **Requests List**: Select the generic list to store access requests.
- **Column for File ID**: Select a Text or Number column in the Requests List to store the ID of the requested file.
- **Column for User Email**: Select a Text column in the Requests List to store the requester's email address.
- **Column for Download Count**: (Optional) Select a Number column to track how many times a user has requested a file. If configured, the system will increment existing requests instead of creating new duplicates.

## Prerequisites

- **SharePoint Online** Tenant.
- **Source Document Library**: A library with columns for Category and Sub-Category.
- **Request Tracking List**: A generic list with columns for File ID, Email, and optionally Count.

## Minimal Path to Awesome

- Clone this repository
- Ensure that you are at the solution folder
- in the command-line run:
  - `npm install`
  - `gulp serve`

## Version history

| Version | Date             | Comments        |
| ------- | ---------------- | --------------- |
| 1.0     | 2025-12-31       | Initial Documentation Update |

## Disclaimer

**THIS CODE IS PROVIDED _AS IS_ WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**