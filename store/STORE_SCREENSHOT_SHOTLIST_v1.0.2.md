# Microsoft Store Screenshot Shot List

Product: Master Budget Automation Tool v1.0.2
Platform: Desktop
Audience: Microsoft Store listing

## Microsoft requirement snapshot

- Format: PNG
- Desktop size: 1366 x 768 or larger
- File size: 50 MB or smaller each
- Minimum required: 1 screenshot
- Recommended: at least 4 screenshots
- Max desktop screenshots: 10
- Optional caption: 200 characters or fewer

Source:
- https://learn.microsoft.com/en-us/windows/apps/publish/publish-your-app/msix/screenshots-and-images
- https://learn.microsoft.com/en-us/windows/apps/publish/publish-your-app/msix/create-app-store-listing

## Recommended screenshot set

### 1. Main import screen

Goal:
Show the overall app layout and its simple three-file workflow.

Capture:
- App open on the main screen
- Expense Sub-Program file selected
- Master Budget template selected
- Output workbook path selected
- Buttons visible: Generate, Create suggested output name, Open output folder, Instructions, Clear

Keep visible:
- Product title
- Three file fields
- Primary action button

Suggested caption:
Import an Expense Sub-Program file into the Master Budget workbook in a simple guided workflow.

### 2. Successful run result

Goal:
Show that the app produces a completed output workbook and a readable run summary.

Capture:
- Progress bar at 100%
- Success banner
- Run summary showing matched rows and matched cells
- Output workbook path visible if possible

Keep visible:
- Green completion state
- Summary panel

Suggested caption:
Review a clear run summary after the workbook is generated.

### 3. Mismatch review

Goal:
Show the audit value of the tool rather than only the happy path.

Capture:
- Run summary with mismatch sections expanded
- Examples of source-only or missing codes
- Warning or highlighted sections visible

Keep visible:
- Difference categories
- Review-friendly status text

Suggested caption:
Spot missing or source-only account and sub-program items before final review.

### 4. Suggested output naming

Goal:
Show that the tool preserves the original template and creates a timestamped output file.

Capture:
- Main screen after clicking Create suggested output name
- Output workbook field showing `_AUTO_YYYYMMDD_HHMM`
- Template and output file paths both visible

Keep visible:
- Different template and output paths
- Timestamped output name

Suggested caption:
Create a timestamped output workbook while keeping the original template unchanged.

### 5. Instructions window

Goal:
Show that the app includes built-in guidance for end users.

Capture:
- Instructions window open
- At least the “How to use the app” section visible
- If helpful, include one embedded Compass screenshot

Keep visible:
- User guidance
- Scrollable instruction content

Suggested caption:
Built-in instructions help users find the correct Compass export and run the import correctly.

## Visual guidance

- Use a clean Windows desktop with no unrelated windows visible.
- Prefer light mode if that is your normal deployment environment.
- Avoid personal email addresses, private file names, or sensitive budget data.
- Replace real finance data with sanitized sample data if needed.
- Keep important text and controls in the top two-thirds of the image.
- Do not add marketing badges, arrows, or extra text overlays inside the screenshot.

## Fast capture plan

1. Prepare one clean sample source file and one clean template.
2. Capture the main screen before running.
3. Capture a successful completed run.
4. Capture a mismatch example using a sample file with intentional differences.
5. Capture the suggested output name state.
6. Capture the Instructions window.

## Suggested final upload order

1. Main import screen
2. Successful run result
3. Mismatch review
4. Suggested output naming
5. Instructions window
