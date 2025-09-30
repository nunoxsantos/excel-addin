# Bill.com Excel Add-in

An Excel Add-in that integrates with the Bill.com API to fetch and display bill data with advanced data transformation capabilities.

## Features

- 🔗 **Bill.com API Integration**: Direct connection to Bill.com's staging API
- 📊 **Data Transformation**: Advanced data processing with customizable rules
- 📈 **Excel Integration**: Seamless data export to Excel spreadsheets
- ⚡ **Real-time Processing**: Live data fetching with pagination support
- 🎯 **Smart Alerts**: Automatic highlighting of high-value and overdue bills
- 🔧 **Customizable Rules**: Easy-to-modify transformation logic

## Prerequisites

- Microsoft Excel (Desktop or Online)
- Node.js (v14 or higher)
- Bill.com API credentials (Session ID and Developer Key)

## Installation

1. **Clone the repository**
   ```bash
   git clone https://github.com/nunoxsantos/excel-addin.git
   cd excel-addin
   ```

2. **Install dependencies**
   ```bash
   npm install
   ```

3. **Build the project**
   ```bash
   npm run build
   ```

4. **Start the development server**
   ```bash
   npm run dev-server
   ```

## Usage

### Development Mode

1. Start the development server:
   ```bash
   npm run dev-server
   ```

2. Open Excel and load the add-in using the manifest file:
   - In Excel, go to **Insert** > **Office Add-ins**
   - Click **Upload My Add-in**
   - Select the `manifest.xml` file from this project

3. Click the **Bill.com Data** button in the Excel ribbon to open the task pane

4. Enter your Bill.com credentials:
   - **Session ID**: Your Bill.com session identifier
   - **Developer Key**: Your Bill.com developer API key

5. Click **Fetch Bills** to retrieve and process your bill data

### Production Deployment

For production deployment, you'll need to:

1. Update the URLs in `manifest.xml` to point to your production server
2. Deploy the built files to a web server
3. Update the `AppDomains` in the manifest to include your production domain

## Data Transformation Rules

The add-in includes several built-in transformation rules:

- **Vendor Name Formatting**: Converts vendor names to uppercase
- **Currency Formatting**: Rounds amounts to 2 decimal places
- **Date Formatting**: Converts dates to readable format
- **High Value Alerts**: Highlights bills over $1,000
- **Overdue Alerts**: Highlights bills past their due date

### Customizing Transformation Rules

You can modify the transformation rules in `src/taskpane/taskpane.ts`:

```typescript
const transformationRules: TransformationRule[] = [
  {
    name: "Your Custom Rule",
    condition: (bill) => /* your condition */,
    transform: (bill) => ({ /* your transformation */ })
  }
];
```

## API Configuration

The add-in is configured to use Bill.com's staging environment:
- **Base URL**: `https://gateway.stage.bill.com/connect/v3/bills`
- **Pagination**: Supports up to 10 pages (200 records max)
- **Rate Limiting**: Built-in safety limits to prevent API overload

## Project Structure

```
excel-addin-mvp/
├── src/
│   ├── taskpane/
│   │   ├── taskpane.html      # Main UI
│   │   ├── taskpane.css       # Styles
│   │   └── taskpane.ts        # Main logic
│   └── commands/
│       ├── commands.html      # Command UI
│       └── commands.ts        # Command logic
├── assets/                    # Icons and images
├── manifest.xml              # Office Add-in manifest
├── package.json              # Dependencies and scripts
└── webpack.config.js         # Build configuration
```

## Available Scripts

- `npm run build` - Build for production
- `npm run build:dev` - Build for development
- `npm run dev-server` - Start development server
- `npm run start` - Start debugging session
- `npm run stop` - Stop debugging session
- `npm run validate` - Validate manifest file
- `npm run lint` - Run ESLint
- `npm run lint:fix` - Fix ESLint issues

## Security Considerations

- **API Keys**: Never commit API keys to version control
- **HTTPS**: Always use HTTPS in production
- **CORS**: Ensure proper CORS configuration for your domain
- **Authentication**: Consider implementing OAuth for production use

## Troubleshooting

### Common Issues

1. **Add-in not loading**: Check that the manifest.xml is valid and URLs are correct
2. **API errors**: Verify your Bill.com credentials and API access
3. **CORS errors**: Ensure your domain is properly configured in Bill.com
4. **Build errors**: Check Node.js version and run `npm install` again

### Debug Mode

Enable debug mode by opening the browser developer tools in Excel:
1. Right-click in the task pane
2. Select "Inspect Element"
3. Check the Console tab for error messages

## Contributing

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/amazing-feature`)
3. Commit your changes (`git commit -m 'Add some amazing feature'`)
4. Push to the branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Support

For support and questions:
- Create an issue on [GitHub](https://github.com/nunoxsantos/excel-addin/issues)
- Check the [Bill.com API documentation](https://developer.bill.com/)

## Changelog

### v1.0.0
- Initial release
- Bill.com API integration
- Data transformation capabilities
- Excel data export
- Pagination support
- Smart alerts for high-value and overdue bills
