# DV News – Technical Documentation

## 1. Overview

DV News is a SharePoint Framework (SPFx) client-side web part that displays promoted SharePoint news items in a responsive carousel. It supports:

- fetching promoted Site Pages from selected SharePoint sites
- displaying news cards with image, title, description, and a read-more link
- showing a modal with detailed item metadata
- allowing admins to create new news items from the UI
- applying department-based visibility and sharing behavior

## 2. Technology Stack

- SPFx 1.22.2
- React 17 with TypeScript
- PnPjs for SharePoint and Microsoft Graph
- Fluent UI React for form controls
- Swiper for the carousel UI
- Moment.js for date formatting
- Heft for build and packaging

## 3. Solution Structure

- [src/webparts/dvNews/DvNewsWebPart.ts](src/webparts/dvNews/DvNewsWebPart.ts) – main web part class and property pane configuration
- [src/webparts/dvNews/components/DvNews.tsx](src/webparts/dvNews/components/DvNews.tsx) – UI, state management, modal dialogs, and carousel rendering
- [src/webparts/dvNews/components/IDvNewsProps.ts](src/webparts/dvNews/components/IDvNewsProps.ts) – props passed from the web part to the React component
- [src/service/SharepointServiceProvider.ts](src/service/SharepointServiceProvider.ts) – service layer for SharePoint/Graph operations
- [src/service/IServiceProvider.ts](src/service/IServiceProvider.ts) – abstraction for the service provider
- [src/service/pnpjsConfig.ts](src/service/pnpjsConfig.ts) – PnPjs initialization and context setup
- [config/](config/) – SPFx configuration and packaging files

## 4. Key Functionalities

The web part provides the following core capabilities:

- Displays promoted news items from SharePoint Site Pages in a responsive carousel.
- Shows a title, date, description, and image for each news item.
- Opens a modal popup with full news details such as title, modified date, author, department, and description.
- Supports an admin-only add-news experience with a dialog form.
- Validates the title before submission and warns users about disallowed special characters.
- Supports department-based filtering and department-aware sharing behavior.
- Allows navigation to the source news page through a Read More link.

## 5. Architecture

The web part follows a simple layered architecture:

1. Web Part Layer
   - The web part initializes the service provider and exposes configurable properties.
   - It renders the React component into the SharePoint page.

2. Presentation Layer
   - The React component manages UI state such as news items, dialogs, loading flags, and form values.
   - It renders the news slider, detail modal, and add-news modal.

3. Service Layer
   - The service provider communicates with SharePoint lists and Microsoft Graph.
   - It retrieves news items, department options, and creates new pages.

## 6. Key Runtime Flow

### 6.1 News Loading

On initial load, the component decides which data source to use based on the selected layout:

- Home layout: fetches news items meant for home display
- Department layout: fetches promoted news items and filters them by department visibility rules

The component uses the service provider to fetch the first page of results and then loads additional items in the background.

### 6.2 Detail Modal

When a user clicks a news title, the component opens a modal and displays:

- title
- modified date
- modified person
- departments
- description

### 6.3 Add News Flow

Admins can open an add-news dialog and submit a new title with department selection. The flow includes:

- title validation
- department selection
- page creation in SharePoint
- department metadata update
- permission sharing for selected departments

## 7. Data Model

The main shape used by the UI is represented by the interface in [src/service/IObject.ts](src/service/IObject.ts):

- Title
- date / Created
- Description
- Image
- Link
- Department
- ShowNewsinHome
- Author / AuthorByLine

## 8. Web Part Properties

The web part exposes the following properties in the property pane:

| Property        | Type          | Description                                                                   |
| --------------- | ------------- | ----------------------------------------------------------------------------- |
| Web Part Title  | Text          | Displays the title shown at the top of the web part.                          |
| Search Sites    | Site Picker   | Allows selection of one or more SharePoint sites to search for promoted news. |
| Target Audience | People Picker | Defines which users are treated as administrators for adding new news items.  |
| Layout          | Dropdown      | Chooses the display mode: Home or Department.                                 |
| Site Title      | Text          | Provides the current site title used for filtering and creation logic.        |

These values are passed into the React component through [src/webparts/dvNews/components/IDvNewsProps.ts](src/webparts/dvNews/components/IDvNewsProps.ts).

## 9. SharePoint Integration Details

The service layer uses PnPjs and SharePoint REST APIs to interact with the Site Pages list.

### Main operations

- getNewsPost – fetches promoted news posts and filters them by department logic
- getNewsPostWithNoDepartment – fetches items for home-style scenarios
- getDepartmentFieldOptions – reads Department field choices from the Site Pages list
- createNewsPost – creates a new client-side page, updates metadata, and shares it with department groups

## 10. Build and Run

From the project root:

```bash
npm install
npm run build
npm start
```

### Scripts

- npm run build – runs tests, cleans, and packages the solution
- npm start – starts the local SPFx debug experience
- npm run clean – removes build output

## 11. Deployment Notes

To deploy the solution:

1. Build the package using the SPFx build pipeline.
2. Package the solution for deployment.
3. Upload the generated package to the tenant or site collection app catalog.
4. Install and trust the app on the target SharePoint site.

## 12. Known Considerations

- The current implementation depends on a Site Pages list with a Department column.
- Some logic is hard-coded to the current SharePoint site path and may need adjustment for other environments.
- The service layer uses console logging during development and could be cleaned up for production.
- The add-news flow relies on department group names matching SharePoint/Graph group display names.

## 13. Future Enhancements

Potential improvements include:

- better error handling and user feedback for create/update operations
- paging or lazy loading improvements for large news sets
- configurable news source list or library name
- richer validation and role-based permission handling
- unit tests for the service layer and UI logic
