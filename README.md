# Export Marketing CRM

Web-based CRM and lead intelligence system for manufacturing exporters selling flanges, fittings, valves, pipes, and related industrial products.

## Included modules

- Lead Database with duplicate prevention
- Contact History logging
- Quotation Management
- Customer Status Pipeline
- Reporting Dashboard
- Reminder System
- Lead Discovery Module
- AI Assistant for email and summary drafts
- Excel export

## Run locally

1. Open a terminal in `crm-app`
2. Run `npm install`
3. Run `npm run dev`
4. Open the local URL shown by Vite

## Production build

- Run `npm run build`
- Open `dist/index.html` or serve the `dist` folder with any static server

## Notes

- Data is stored in browser `localStorage`
- Sample export-marketing data loads automatically on first open
- Excel export downloads a workbook with leads, contacts, quotations, and reminders
