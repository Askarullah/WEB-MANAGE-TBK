# WEB-MANAGE-TBK

A comprehensive Flask web application for managing and monitoring MikroTik network devices. This application provides a user-friendly interface for tracking network equipment, managing device configurations, and monitoring network status across multiple branches.

## Features

### 🌐 Network Device Management
- **Device Tracking**: Monitor active and suspended network devices
- **Real-time Status Checking**: Check device connectivity and status
- **Batch Operations**: Perform bulk operations on multiple devices
- **Device Mapping**: CO (Central Office) to device mapping

### 🔧 MikroTik Integration
- **SSH Connection Management**: Secure connections to MikroTik routers
- **Firewall Management**: View and manage firewall rules
- **Queue Management**: Simple queue parser and management
- **Multi-device Support**: Support for multiple MikroTik devices with different credentials

### 📊 Data Management
- **Excel Integration**: Upload and process Excel files
- **Data Export**: Export data in various formats
- **Search Functionality**: Search across ODP, IP addresses, and CO mappings
- **Batch Processing**: Handle multiple data entries efficiently

### 🖥️ Web Interface
- **Responsive Design**: Modern, mobile-friendly interface
- **Multiple Views**: 
  - Home dashboard
  - Active device management
  - Tracking interfaces (CO, ODP, IP)
  - Automation tools
  - Queue management

## Installation

### Prerequisites
- Python 3.7 or higher
- pip (Python package installer)

### Setup

1. **Clone the repository**
   ```bash
   git clone <repository-url>
   cd WEB-MANAGE-TBK
   ```

2. **Install dependencies**
   ```bash
   pip install -r requirements.txt
   ```

3. **Configure environment variables**
   
   Copy the `.env` file and update the configuration:
   ```bash
   cp .env.example .env
   ```
   
   Edit `.env` with your settings:
   ```env
   # Flask Configuration
   FLASK_SECRET_KEY=your-secret-key-here
   FLASK_HOST=localhost
   FLASK_PORT=5050
   FLASK_DEBUG=False
   
   # MikroTik SSH Configuration
   MIKROTIK_USERNAME=your-username
   MIKROTIK_PASSWORD=your-password
   MIKROTIK_SSH_PORT=22
   
   # Additional MikroTik devices (up to 10)
   MIKROTIK_USERNAME_2=username2
   MIKROTIK_PASSWORD_2=password2
   # ... and so on
   ```

4. **Prepare device configuration files**
   
   Ensure you have the following JSON files in the project root:
   - `devices-active.json` - Active network devices
   - `devices-suspend.json` - Suspended network devices

## Usage

### Starting the Application

```bash
python script-web-track.py
```

The application will start on `http://localhost:5050` (or your configured host/port).

### Web Interface Navigation

1. **Home Dashboard** (`/`) - Main overview and navigation
2. **Manage Active** (`/manage-active.html`) - Manage active devices
3. **Tracking CO** (`/tracking-co.html`) - Central Office tracking
4. **Tracking ODP** (`/tracking-odp.html`) - ODP (Optical Distribution Point) tracking
5. **Tracking IP** (`/tracking-ip.html`) - IP address tracking
6. **Automate Suspend** (`/automate-suspend.html`) - Device suspension automation
7. **Queue Parser** (`/simple-queue-parser.html`) - MikroTik queue management

### API Endpoints

The application provides several REST API endpoints:

- `POST /api/mikrotik/test-connection` - Test MikroTik connection
- `POST /api/mikrotik/firewall-list` - Get firewall rules
- `POST /api/mikrotik/batch-status-check` - Batch device status check
- `GET /api/devices-active` - Get active devices
- `GET /api/devices-suspend` - Get suspended devices
- `POST /upload` - Upload Excel files
- `POST /export` - Export data

## Configuration

### Device Configuration

Devices are configured in JSON format:

```json
[
  {
    "name": "CO-AZAM-R1-250",
    "ip": "103.248.217.250",
    "co": "CO-AZAM-R1",
    "location": "AZAM Branch",
    "description": "CO Azam"
  }
]
```

### MikroTik Configuration

The application supports multiple MikroTik devices with different credentials. Configure them in the `.env` file using the pattern:

```env
MIKROTIK_USERNAME_X=username
MIKROTIK_PASSWORD_X=password
```

Where X is a number from 1 to 10.

## Dependencies

- **Flask 2.3.3** - Web framework
- **flask-cors 4.0.0** - Cross-Origin Resource Sharing
- **pandas 2.1.1** - Data manipulation and analysis
- **openpyxl 3.1.2** - Excel file handling
- **paramiko 3.3.1** - SSH client for MikroTik connections
- **python-dotenv 1.0.0** - Environment variable management
- **Werkzeug 2.3.7** - WSGI utilities

## Security Considerations

1. **Change the default secret key** in production
2. **Use strong passwords** for MikroTik devices
3. **Limit network access** to the application
4. **Keep credentials secure** in the `.env` file
5. **Regular updates** of dependencies

## File Structure

```
WEB-MANAGE-TBK/
├── script-web-track.py          # Main Flask application
├── requirements.txt             # Python dependencies
├── .env                        # Environment configuration
├── devices-active.json         # Active devices configuration
├── devices-suspend.json        # Suspended devices configuration
├── templates/                  # HTML templates
│   ├── home.html
│   ├── manage-active.html
│   ├── tracking-co.html
│   ├── tracking-odp.html
│   ├── tracking-ip.html
│   ├── automate-suspend.html
│   └── simple-queue-parser.html
└── README.md                   # This file
```

## Troubleshooting

### Common Issues

1. **Connection refused to MikroTik devices**
   - Check SSH is enabled on MikroTik devices
   - Verify credentials in `.env` file
   - Ensure network connectivity

2. **File upload errors**
   - Check file size limits (default 50MB)
   - Ensure proper Excel file format

3. **Application won't start**
   - Verify all dependencies are installed
   - Check port availability
   - Review `.env` configuration

### Logs and Debugging

The application includes debug logging. Check the console output for error messages and connection status.

## Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Test thoroughly
5. Submit a pull request

## License

This project is licensed under the MIT License - see the LICENSE file for details.

## Support

For support and questions, please create an issue in the repository or contact the development team.

---

**Note**: This application is designed for network administrators managing MikroTik infrastructure. Ensure proper network security measures are in place when deploying in production environments.