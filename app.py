# app.py
from flask import Flask, request, render_template, redirect, url_for, flash, jsonify, send_file
from flask_pymongo import PyMongo
from bson.objectid import ObjectId
import pandas as pd
from werkzeug.utils import secure_filename
import os
import json
import uuid
import datetime
from pathlib import Path
import io
import gridfs
from flask_restx import Api, Resource, fields
import logging
from openpyxl import load_workbook, Workbook
import calendar
import pymongo.errors

# Set up logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler("app.log"),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

# Initialize Flask app
app = Flask(__name__)
# Change this to a secure random key in production
app.secret_key = "your-secret-key"
app.config["UPLOAD_FOLDER"] = "uploads"
app.config["ALLOWED_EXTENSIONS"] = {"xlsx", "xls"}
app.config["MONGO_URI"] = "mongodb://localhost:27017/userRoster"
app.config["MAX_CONTENT_LENGTH"] = 20 * 1024 * 1024  # 20MB max file size

# Initialize Swagger documentation
api = Api(app,
          version='1.0',
          title='User Roster API',
          description='API for managing user roster data from Excel files',
          doc='/api/docs',
          prefix='/api')

# Create namespaces for API documentation
ns_upload = api.namespace('uploads', description='Upload operations')
ns_users = api.namespace('users', description='User operations')
ns_roster = api.namespace('roster', description='Roster operations')

# Define API models for Swagger documentation
upload_model = api.model('Upload', {
    'filename': fields.String(description='Unique filename for the uploaded file'),
    'original_filename': fields.String(description='Original filename'),
    'upload_date': fields.DateTime(description='Date and time of upload'),
    'file_size': fields.Integer(description='Size of the file in bytes'),
    'record_count': fields.Integer(description='Number of records in the file'),
    'month': fields.Integer(description='Month of the roster (1-12)'),
    'year': fields.Integer(description='Year of the roster')
})

user_model = api.model('User', {
    '_id': fields.String(description='User ID'),
    'upload_id': fields.String(description='ID of the upload that created this user'),
    'month': fields.Integer(description='Month of the roster (1-12)'),
    'year': fields.Integer(description='Year of the roster')
    # Note: Other fields will be dynamically determined from Excel file
})

upload_response = api.model('UploadResponse', {
    'success': fields.Boolean(description='Success status'),
    'message': fields.String(description='Response message'),
    'filename': fields.String(description='Unique filename for the uploaded file'),
    'upload_id': fields.String(description='ID of the created upload'),
    'redirect': fields.String(description='Redirection URL')
})

roster_input = api.model('RosterInput', {
    'month': fields.Integer(required=True, description='Month of the roster (1-12)'),
    'year': fields.Integer(required=True, description='Year of the roster'),
    'employee_name': fields.String(required=True, description='Employee name'),
    'employee_id': fields.String(required=True, description='Employee ID'),
    'department': fields.String(required=True, description='Department'),
    'shifts': fields.Raw(required=True, description='Dictionary of day numbers to shift codes')
})

# Initialize MongoDB connection
try:
    mongo = PyMongo(app)
    # Setup GridFS for file storage
    fs = gridfs.GridFS(mongo.db)
    logger.info("Successfully connected to MongoDB")
except Exception as e:
    logger.error(f"Failed to connect to MongoDB: {str(e)}")
    raise


def initialize_database():
    """Initialize database structure if it doesn't exist"""
    try:
        # Check if collections exist and create them if they don't
        collection_names = mongo.db.list_collection_names()

        # Create users collection if it doesn't exist
        if 'users' not in collection_names:
            logger.info("Creating users collection")
            mongo.db.create_collection('users')
            # Create indexes for users collection
            mongo.db.users.create_index([('month', 1), ('year', 1)])
            mongo.db.users.create_index([('upload_id', 1)])
            mongo.db.users.create_index([('employee_id', 1)])

        # Create uploads collection if it doesn't exist
        if 'uploads' not in collection_names:
            logger.info("Creating uploads collection")
            mongo.db.create_collection('uploads')
            # Create indexes for uploads collection
            mongo.db.uploads.create_index([('month', 1), ('year', 1)])
            mongo.db.uploads.create_index([('upload_date', -1)])

        logger.info("Database initialization completed successfully")
        return True
    except Exception as e:
        logger.error(f"Error initializing database: {str(e)}")
        return False


# Initialize database on startup
initialization_status = initialize_database()
if not initialization_status:
    logger.warning(
        "Database initialization failed. Some features may not work correctly.")

# Create uploads folder if it doesn't exist (for temporary storage)
try:
    os.makedirs(app.config["UPLOAD_FOLDER"], exist_ok=True)
    logger.info(
        f"Upload folder created at {os.path.abspath(app.config['UPLOAD_FOLDER'])}")
except Exception as e:
    logger.error(f"Failed to create upload folder: {str(e)}")
    raise


def allowed_file(filename):
    """Check if the uploaded file has an allowed extension"""
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in app.config["ALLOWED_EXTENSIONS"]


def process_excel_with_merged_cells(filepath):
    """
    Process Excel file with merged cells by unmerging and filling values

    Args:
        filepath: Path to the Excel file

    Returns:
        Path to processed Excel file
    """
    logger.info(f"Processing Excel file with merged cells: {filepath}")
    workbook = load_workbook(filepath)

    # Process each sheet
    for sheet_name in workbook.sheetnames:
        sheet = workbook[sheet_name]

        # Get all merged cell ranges
        merged_ranges = list(sheet.merged_cells.ranges)
        logger.info(
            f"Sheet '{sheet_name}' has {len(merged_ranges)} merged ranges")

        # Store values from merged cells before unmerging
        merged_values = {}
        for merged_range in merged_ranges:
            # Get the value from the top-left cell
            top_left_cell = sheet.cell(
                merged_range.min_row, merged_range.min_col)
            value = top_left_cell.value
            merged_values[merged_range] = value

        # Unmerge all cells
        for merged_range in merged_ranges:
            sheet.unmerge_cells(range_string=str(merged_range))

        # Fill in values for the previously merged cells
        for merged_range, value in merged_values.items():
            for row in range(merged_range.min_row, merged_range.max_row + 1):
                for col in range(merged_range.min_col, merged_range.max_col + 1):
                    sheet.cell(row=row, column=col).value = value

    # Save to a new file
    processed_filepath = f"{os.path.splitext(filepath)[0]}_processed.xlsx"
    workbook.save(processed_filepath)
    logger.info(f"Saved processed Excel file to: {processed_filepath}")

    return processed_filepath


def create_empty_roster_template(month, year):
    """
    Create an empty roster template Excel file for the given month and year

    Args:
        month: Month number (1-12)
        year: Year

    Returns:
        Path to the created Excel file
    """
    # Create a new workbook
    wb = Workbook()
    ws = wb.active
    ws.title = f"{calendar.month_name[month]} {year}"

    # Get the number of days in the month
    num_days = calendar.monthrange(year, month)[1]

    # Create header row
    ws.cell(row=1, column=1).value = f"{calendar.month_name[month]}"
    ws.cell(row=1, column=2).value = "Employee Name"
    ws.cell(row=1, column=3).value = "Employee ID"
    ws.cell(row=1, column=4).value = "Department"

    # Add day columns
    for day in range(1, num_days + 1):
        ws.cell(row=1, column=4 + day).value = day

    # Add some empty rows for data entry
    for row in range(2, 12):
        ws.cell(
            row=row, column=1).value = f"{calendar.month_name[month]} {year}"
        for col in range(2, 5 + num_days):
            ws.cell(row=row, column=col).value = ""

    # Save the file
    filepath = os.path.join(
        app.config['UPLOAD_FOLDER'], f"roster_template_{month}_{year}.xlsx")
    wb.save(filepath)

    return filepath


def overwrite_roster_for_month_year(month, year, records):
    """
    Overwrite all roster records for a specific month and year

    Args:
        month: Month number (1-12)
        year: Year
        records: New records to insert

    Returns:
        Number of deleted records
    """
    try:
        # Delete existing records
        result = mongo.db.users.delete_many({'month': month, 'year': year})
        deleted_count = result.deleted_count
        logger.info(
            f"Deleted {deleted_count} existing records for {month}/{year}")

        # Add month and year to each record if not already present
        for record in records:
            record['month'] = month
            record['year'] = year

            # Ensure no _id field is explicitly set to avoid duplicate key errors
            if '_id' in record:
                del record['_id']

        # Insert new records
        if records:
            try:
                mongo.db.users.insert_many(records)
                logger.info(
                    f"Inserted {len(records)} new records for {month}/{year}")
            except pymongo.errors.BulkWriteError as bwe:
                logger.error(f"Bulk write error: {bwe.details}")
                # Continue with other operations - records may have been partially inserted

        return deleted_count
    except Exception as e:
        logger.error(f"Error overwriting roster for {month}/{year}: {str(e)}")
        raise

# ----------------------------
# Web Routes (HTML UI)
# ----------------------------


@app.route('/')
def index():
    """Render the main page with month/year selection and upload form"""
    try:
        # Check if database is initialized
        if not mongo.db.list_collection_names():
            initialize_database()

        # Generate month choices
        current_month = datetime.datetime.now().month
        current_year = datetime.datetime.now().year

        months = [(i, calendar.month_name[i]) for i in range(1, 13)]
        years = [(y, y) for y in range(current_year - 2, current_year + 3)]

        return render_template('index.html',
                               months=months,
                               years=years,
                               current_month=current_month,
                               current_year=current_year)
    except Exception as e:
        logger.error(f"Error rendering index page: {str(e)}")
        return render_template('error.html', error=str(e))


@app.route('/roster')
@app.route('/roster/<int:month>/<int:year>')
def roster(month=None, year=None):
    """Display the roster data for specific month and year"""
    try:
        # Check if collections exist
        if 'users' not in mongo.db.list_collection_names() or 'uploads' not in mongo.db.list_collection_names():
            initialize_database()
            flash("Database is being initialized. Please upload a roster file.")
            return redirect(url_for('index'))

        # If month/year not provided, use current
        if month is None or year is None:
            current_date = datetime.datetime.now()
            month = current_date.month
            year = current_date.year

        # Query for users in this month/year
        query = {'month': month, 'year': year}
        try:
            users = list(mongo.db.users.find(query).sort('employee_name', 1))
        except Exception as e:
            logger.error(f"Error querying users: {str(e)}")
            users = []
            flash("Error retrieving user data. Please try again.")

        # Get all months/years with data for the filter dropdown
        try:
            pipeline = [
                {'$group': {'_id': {'month': '$month', 'year': '$year'}}},
                {'$sort': {'_id.year': -1, '_id.month': -1}}
            ]
            available_periods = list(mongo.db.users.aggregate(pipeline))
        except Exception as e:
            logger.error(f"Error retrieving available periods: {str(e)}")
            available_periods = []

        # Format periods for display
        formatted_periods = []
        for period in available_periods:
            if period['_id'] and 'month' in period['_id'] and 'year' in period['_id']:
                m = period['_id']['month']
                y = period['_id']['year']
                if m and y and 1 <= m <= 12:  # Validate month value
                    formatted_periods.append({
                        'month': m,
                        'year': y,
                        'name': f"{calendar.month_name[m]} {y}",
                        'selected': m == month and y == year
                    })

        # Get the days in month for column headers
        days_in_month = calendar.monthrange(year, month)[1]
        day_columns = list(range(1, days_in_month + 1))

        # Convert ObjectId to string for JSON serialization
        for user in users:
            user['_id'] = str(user['_id'])

        return render_template(
            'roster.html',
            users=users,
            periods=formatted_periods,
            current_month=month,
            current_year=year,
            month_name=calendar.month_name[month],
            day_columns=day_columns
        )
    except Exception as e:
        logger.error(f"Error displaying roster: {str(e)}")
        flash(f"An error occurred: {str(e)}")
        return render_template('error.html', error=str(e))


@app.route('/uploads')
def list_uploads():
    """List all uploads"""
    try:
        # Check if collections exist
        if 'uploads' not in mongo.db.list_collection_names():
            initialize_database()
            flash("Database is being initialized. No uploads found.")
            uploads = []
        else:
            uploads = list(mongo.db.uploads.find().sort('upload_date', -1))

            # Convert ObjectIds to strings for JSON serialization
            for upload in uploads:
                upload['_id'] = str(upload['_id'])
                upload['file_id'] = str(upload['file_id'])
                # Add month name
                if 'month' in upload and upload['month'] and 1 <= upload['month'] <= 12:
                    upload['month_name'] = calendar.month_name[upload['month']]

        return render_template('uploads.html', uploads=uploads)
    except Exception as e:
        logger.error(f"Error listing uploads: {str(e)}")
        flash(f"An error occurred: {str(e)}")
        return render_template('error.html', error=str(e))


@app.route('/download/<file_id>')
def download_file(file_id):
    """Download a file from GridFS"""
    try:
        # Get file from GridFS
        grid_out = fs.get(ObjectId(file_id))

        # Create file-like object
        file_data = io.BytesIO(grid_out.read())

        # Reset file pointer
        file_data.seek(0)

        # Return file
        return send_file(
            file_data,
            mimetype=grid_out.content_type,
            download_name=grid_out.original_filename,
            as_attachment=True
        )
    except Exception as e:
        logger.error(f"Error downloading file: {str(e)}")
        flash(f"Error downloading file: {str(e)}")
        return redirect(url_for('roster'))


@app.route('/download-template/<int:month>/<int:year>')
def download_template(month, year):
    """Generate and download an empty roster template"""
    try:
        # Create template
        template_path = create_empty_roster_template(month, year)

        # Send file
        return send_file(
            template_path,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            download_name=f"roster_template_{calendar.month_name[month]}_{year}.xlsx",
            as_attachment=True
        )
    except Exception as e:
        logger.error(f"Error creating template: {str(e)}")
        flash(f"Error creating template: {str(e)}")
        return redirect(url_for('index'))


@app.route('/manual-entry')
def manual_entry_form():
    """Show form for manual roster entry"""
    try:
        # Get current month and year
        current_date = datetime.datetime.now()
        month = current_date.month
        year = current_date.year

        # Generate month choices
        months = [(i, calendar.month_name[i]) for i in range(1, 13)]
        years = [(y, y)
                 for y in range(current_date.year - 2, current_date.year + 3)]

        # Get days in the current month
        days_in_month = calendar.monthrange(year, month)[1]
        day_columns = list(range(1, days_in_month + 1))

        # Define shift codes
        shift_codes = [
            {'code': 'G', 'name': 'General (9-5)', 'color': 'bg-gray-200'},
            {'code': 'S1', 'name': 'Shift 1 (7-3)', 'color': 'bg-blue-200'},
            {'code': 'S2', 'name': 'Shift 2 (3-11)', 'color': 'bg-green-200'},
            {'code': 'S3', 'name': 'Shift 3 (11-7)', 'color': 'bg-yellow-200'},
            {'code': 'W', 'name': 'Weekend/Off', 'color': 'bg-red-200'},
            {'code': 'V', 'name': 'Vacation', 'color': 'bg-purple-200'},
            {'code': 'T', 'name': 'Training', 'color': 'bg-indigo-200'},
            {'code': 'O', 'name': 'On-Call', 'color': 'bg-pink-200'},
        ]

        return render_template(
            'manual_entry.html',
            months=months,
            years=years,
            current_month=month,
            current_year=year,
            month_name=calendar.month_name[month],
            day_columns=day_columns,
            shift_codes=shift_codes
        )
    except Exception as e:
        logger.error(f"Error showing manual entry form: {str(e)}")
        flash(f"An error occurred: {str(e)}")
        return render_template('error.html', error=str(e))

# ----------------------------
# API Routes (RESTful endpoints)
# ----------------------------


@ns_upload.route('/')
class UploadList(Resource):
    @api.doc('list_uploads')
    @api.marshal_list_with(upload_model)
    def get(self):
        """Get all uploads"""
        try:
            uploads = list(mongo.db.uploads.find().sort('upload_date', -1))

            # Convert ObjectIds to strings for JSON serialization
            for upload in uploads:
                upload['_id'] = str(upload['_id'])
                upload['file_id'] = str(upload['file_id'])

            return uploads
        except Exception as e:
            logger.error(f"API error listing uploads: {str(e)}")
            api.abort(500, f"Server error: {str(e)}")

    @api.doc('create_upload')
    @api.expect(api.parser()
                .add_argument('file', location='files', type='file', required=True, help='Excel file (.xlsx, .xls)')
                .add_argument('month', type=int, required=True, help='Month (1-12)')
                .add_argument('year', type=int, required=True, help='Year'))
    @api.response(201, 'Upload successful')
    @api.response(400, 'Bad request')
    @api.response(500, 'Server error')
    @api.marshal_with(upload_response)
    def post(self):
        """Upload an Excel file for a specific month and year"""
        # Get month and year
        month = request.form.get('month', type=int)
        year = request.form.get('year', type=int)

        # Validate month and year
        if not month or month < 1 or month > 12:
            logger.warning(f"Invalid month: {month}")
            api.abort(400, "Month must be between 1 and 12")

        if not year or year < 2000 or year > 2100:
            logger.warning(f"Invalid year: {year}")
            api.abort(400, "Year must be between 2000 and 2100")

        # Check if there's a file in the request
        if 'file' not in request.files:
            logger.warning("No file part in the request")
            api.abort(400, "No file part in the request")

        file = request.files['file']

        # Check if file is empty
        if file.filename == '':
            logger.warning("No file selected")
            api.abort(400, "No file selected")

        # Check file type
        if not allowed_file(file.filename):
            logger.warning(f"Invalid file type: {file.filename}")
            api.abort(
                400, "File type not allowed. Please upload an Excel file (.xlsx or .xls)")

        # Save and process the file
        temp_filepath = None
        processed_filepath = None

        try:
            # Get original filename and extension
            original_filename = secure_filename(file.filename)
            file_extension = Path(original_filename).suffix
            file_basename = Path(original_filename).stem

            # Create a unique filename
            timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            unique_id = str(uuid.uuid4())[:8]  # Use first 8 chars of UUID

            # New filename format: original_name_YYYYMMDD_HHMMSS_uuid.extension
            unique_filename = f"{file_basename}_{timestamp}_{unique_id}{file_extension}"

            # Temporary local filepath for processing
            temp_filepath = os.path.join(
                app.config['UPLOAD_FOLDER'], unique_filename)
            logger.info(f"Saving temporary file to {temp_filepath}")

            # Save file temporarily to process it
            file.save(temp_filepath)

            # Process the Excel file to handle merged cells
            processed_filepath = process_excel_with_merged_cells(temp_filepath)

            # Read the processed Excel file
            logger.info(f"Reading processed Excel file {processed_filepath}")
            df = pd.read_excel(processed_filepath, engine='openpyxl')

            # Check if data exists
            if df.empty:
                # Delete temporary files
                os.remove(temp_filepath)
                if os.path.exists(processed_filepath):
                    os.remove(processed_filepath)
                logger.warning("The Excel file contains no data")
                api.abort(400, "The Excel file contains no data")

            # Convert DataFrame to list of dictionaries and fix numeric keys
            records = df.to_dict('records')
            logger.info(f"Extracted {len(records)} records from Excel file")

            # Process records to convert numeric keys to strings
            processed_records = []
            for record in records:
                processed_record = {}
                for key, value in record.items():
                    # Convert any non-string keys to strings
                    string_key = str(key) if not isinstance(key, str) else key
                    processed_record[string_key] = value
                processed_records.append(processed_record)

            # Store file in GridFS
            file_id = None
            with open(temp_filepath, 'rb') as f:
                # Store in GridFS with metadata
                file_id = fs.put(
                    f.read(),
                    filename=unique_filename,
                    original_filename=original_filename,
                    content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' if file_extension.lower() == '.xlsx' else 'application/vnd.ms-excel',
                    upload_date=datetime.datetime.now(),
                    month=month,
                    year=year
                )
                logger.info(f"Stored file in GridFS with ID {file_id}")

            # Create upload record with reference to GridFS file
            upload_info = {
                'filename': unique_filename,
                'original_filename': original_filename,
                'upload_date': datetime.datetime.now(),
                'file_size': os.path.getsize(temp_filepath) if os.path.exists(temp_filepath) else 0,
                'file_id': file_id,
                'record_count': len(processed_records),
                'month': month,
                'year': year
            }

            # Insert upload info to MongoDB
            upload_id = mongo.db.uploads.insert_one(upload_info).inserted_id
            logger.info(f"Created upload record with ID {upload_id}")

            # Overwrite existing records for this month/year
            try:
                deleted_count = overwrite_roster_for_month_year(
                    month, year, processed_records)
                logger.info(
                    f"Overwritten {deleted_count} records for {calendar.month_name[month]} {year}")

                # Clean up temporary files
                if os.path.exists(temp_filepath):
                    os.remove(temp_filepath)
                if os.path.exists(processed_filepath):
                    os.remove(processed_filepath)

                overwrote_message = f" (overwrote {deleted_count} existing records)" if deleted_count > 0 else ""

                return {
                    'success': True,
                    'message': f'Successfully uploaded {len(processed_records)} records for {calendar.month_name[month]} {year}{overwrote_message}',
                    'filename': unique_filename,
                    'upload_id': str(upload_id),
                    'redirect': url_for('roster', month=month, year=year)
                }, 201
            except Exception as e:
                logger.error(f"Error during overwrite operation: {str(e)}")
                # Even if overwrite fails, we created the upload record, so return success with a warning
                return {
                    'success': True,
                    'message': f'File uploaded but there was an error updating records: {str(e)}',
                    'filename': unique_filename,
                    'upload_id': str(upload_id),
                    'redirect': url_for('roster', month=month, year=year)
                }, 201

        except Exception as e:
            # Clean up temp files if they exist
            if temp_filepath and os.path.exists(temp_filepath):
                os.remove(temp_filepath)
            if processed_filepath and os.path.exists(processed_filepath):
                os.remove(processed_filepath)

            logger.error(f"Error during file upload: {str(e)}")
            import traceback
            logger.error(traceback.format_exc())

            api.abort(500, f"Error processing file: {str(e)}")


@ns_upload.route('/<upload_id>')
@api.doc(params={'upload_id': 'The upload ID'})
class Upload(Resource):
    @api.doc('get_upload')
    @api.marshal_with(upload_model)
    @api.response(404, 'Upload not found')
    def get(self, upload_id):
        """Get a specific upload"""
        try:
            upload = mongo.db.uploads.find_one({'_id': ObjectId(upload_id)})
            if not upload:
                logger.warning(f"Upload not found: {upload_id}")
                api.abort(404, "Upload not found")

            # Convert ObjectId to string for JSON serialization
            upload['_id'] = str(upload['_id'])
            upload['file_id'] = str(upload['file_id'])

            return upload
        except Exception as e:
            logger.error(f"Error retrieving upload {upload_id}: {str(e)}")
            api.abort(500, f"Server error: {str(e)}")

    @api.doc('delete_upload')
    @api.response(204, 'Upload deleted')
    @api.response(404, 'Upload not found')
    def delete(self, upload_id):
        """Delete an upload and all associated records"""
        try:
            # Find the upload
            upload = mongo.db.uploads.find_one({'_id': ObjectId(upload_id)})
            if not upload:
                logger.warning(f"Upload not found for deletion: {upload_id}")
                api.abort(404, "Upload not found")

            # Delete file from GridFS
            fs.delete(ObjectId(upload['file_id']))
            logger.info(f"Deleted file {upload['file_id']} from GridFS")

            # Delete all users associated with this upload
            result = mongo.db.users.delete_many({'upload_id': upload_id})
            logger.info(
                f"Deleted {result.deleted_count} users associated with upload {upload_id}")

            # Delete the upload record
            mongo.db.uploads.delete_one({'_id': ObjectId(upload_id)})
            logger.info(f"Deleted upload record {upload_id}")

            return '', 204
        except Exception as e:
            logger.error(f"Error deleting upload {upload_id}: {str(e)}")
            api.abort(500, f"Server error: {str(e)}")


@ns_roster.route('/')
class RosterManagement(Resource):
    @api.doc('create_roster_entry')
    @api.expect(roster_input)
    @api.response(201, 'Roster entry created')
    @api.response(400, 'Bad request')
    def post(self):
        """Create a new roster entry manually"""
        try:
            data = request.json

            # Validate required fields
            required_fields = ['month', 'year', 'employee_name',
                               'employee_id', 'department', 'shifts']
            for field in required_fields:
                if field not in data:
                    return {'success': False, 'message': f'Missing required field: {field}'}, 400

            # Validate month and year
            month = data['month']
            year = data['year']

            if month < 1 or month > 12:
                return {'success': False, 'message': 'Month must be between 1 and 12'}, 400

            if year < 2000 or year > 2100:
                return {'success': False, 'message': 'Year must be between 2000 and 2100'}, 400

            # Prepare record
            record = {
                'month': month,
                'year': year,
                'employee_name': data['employee_name'],
                'employee_id': data['employee_id'],
                'department': data['department'],
                'upload_id': None  # Manual entry has no upload
            }

            # Add shifts
            shifts = data['shifts']
            if not isinstance(shifts, dict):
                return {'success': False, 'message': 'Shifts must be a dictionary of day numbers to shift codes'}, 400

            # Get days in month
            days_in_month = calendar.monthrange(year, month)[1]

            # Add each day's shift
            for day in range(1, days_in_month + 1):
                day_str = str(day)
                if day_str in shifts:
                    record[day_str] = shifts[day_str]
                else:
                    record[day_str] = ''  # Empty for days without shifts

            # Insert the record
            result = mongo.db.users.insert_one(record)

            return {
                'success': True,
                'message': 'Roster entry created successfully',
                '_id': str(result.inserted_id)
            }, 201

        except Exception as e:
            logger.error(f"Error creating roster entry: {str(e)}")
            return {'success': False, 'message': f'Error: {str(e)}'}, 500


@ns_roster.route('/<int:month>/<int:year>')
@api.doc(params={'month': 'Month (1-12)', 'year': 'Year'})
class RosterPeriod(Resource):
    @api.doc('get_roster_for_period')
    def get(self, month, year):
        """Get roster data for a specific month and year"""
        try:
            # Validate month and year
            if month < 1 or month > 12:
                return {'success': False, 'message': 'Month must be between 1 and 12'}, 400

            if year < 2000 or year > 2100:
                return {'success': False, 'message': 'Year must be between 2000 and 2100'}, 400

            # Query for users in this month/year
            query = {'month': month, 'year': year}
            users = list(mongo.db.users.find(query))

            # Convert ObjectId to string for JSON serialization
            for user in users:
                user['_id'] = str(user['_id'])
                if 'upload_id' and user['upload_id']:
                    user['upload_id'] = str(user['upload_id'])

            return {
                'success': True,
                'month': month,
                'month_name': calendar.month_name[month],
                'year': year,
                'days_in_month': calendar.monthrange(year, month)[1],
                'record_count': len(users),
                'records': users
            }

        except Exception as e:
            logger.error(
                f"Error retrieving roster for {month}/{year}: {str(e)}")
            return {'success': False, 'message': f'Error: {str(e)}'}, 500


@ns_users.route('/')
class UserList(Resource):
    @api.doc('list_users')
    @api.param('month', 'Filter users by month (1-12)')
    @api.param('year', 'Filter users by year')
    @api.param('upload_id', 'Filter users by upload ID')
    def get(self):
        """Get all users, optionally filtered by month, year, or upload ID"""
        try:
            # Get query parameters
            month = request.args.get('month', type=int)
            year = request.args.get('year', type=int)
            upload_id = request.args.get('upload_id')

            # Build query
            query = {}
            if month:
                query['month'] = month
            if year:
                query['year'] = year
            if upload_id:
                query['upload_id'] = upload_id

            users = list(mongo.db.users.find(query))

            # Convert ObjectId to string for JSON serialization
            for user in users:
                user['_id'] = str(user['_id'])
                if 'upload_id' in user and user['upload_id']:
                    user['upload_id'] = str(user['upload_id'])

            return {'users': users, 'count': len(users)}
        except Exception as e:
            logger.error(f"Error listing users: {str(e)}")
            api.abort(500, f"Server error: {str(e)}")


@ns_users.route('/<user_id>')
@api.doc(params={'user_id': 'The user ID'})
class User(Resource):
    @api.doc('get_user')
    @api.response(404, 'User not found')
    def get(self, user_id):
        """Get a specific user"""
        try:
            user = mongo.db.users.find_one({'_id': ObjectId(user_id)})

            if not user:
                logger.warning(f"User not found: {user_id}")
                api.abort(404, "User not found")

            # Convert ObjectId to string for JSON serialization
            user['_id'] = str(user['_id'])
            if 'upload_id' in user and user['upload_id']:
                user['upload_id'] = str(user['upload_id'])

            return user
        except Exception as e:
            logger.error(f"Error retrieving user {user_id}: {str(e)}")
            api.abort(500, f"Server error: {str(e)}")

    @api.doc('update_user')
    @api.expect(user_model)
    @api.response(200, 'User updated')
    @api.response(404, 'User not found')
    def put(self, user_id):
        """Update a specific user"""
        data = request.json

        # Remove _id from the data if it exists
        if '_id' in data:
            del data['_id']

        try:
            result = mongo.db.users.update_one(
                {'_id': ObjectId(user_id)},
                {'$set': data}
            )

            if result.matched_count == 0:
                logger.warning(f"User not found for update: {user_id}")
                api.abort(404, "User not found")

            logger.info(f"Updated user {user_id}")
            return {'success': True, 'message': 'User updated successfully'}
        except Exception as e:
            logger.error(f"Error updating user {user_id}: {str(e)}")
            api.abort(500, f"Server error: {str(e)}")

    @api.doc('delete_user')
    @api.response(204, 'User deleted')
    @api.response(404, 'User not found')
    def delete(self, user_id):
        """Delete a specific user"""
        try:
            result = mongo.db.users.delete_one({'_id': ObjectId(user_id)})

            if result.deleted_count == 0:
                logger.warning(f"User not found for deletion: {user_id}")
                api.abort(404, "User not found")

            logger.info(f"Deleted user {user_id}")
            return '', 204
        except Exception as e:
            logger.error(f"Error deleting user {user_id}: {str(e)}")
            api.abort(500, f"Server error: {str(e)}")

# ----------------------------
# AJAX Routes (for the web UI)
# ----------------------------


@app.route('/upload', methods=['POST'])
def upload_file():
    """Handle file upload and import to MongoDB with GridFS storage"""
    # Get month and year
    month = request.form.get('month', type=int)
    year = request.form.get('year', type=int)

    # Validate month and year
    if not month or month < 1 or month > 12:
        logger.warning(f"Invalid month: {month}")
        return jsonify({'success': False, 'message': 'Month must be between 1 and 12'}), 400

    if not year or year < 2000 or year > 2100:
        logger.warning(f"Invalid year: {year}")
        return jsonify({'success': False, 'message': 'Year must be between 2000 and 2100'}), 400

    # Check if there's a file in the request
    if 'file' not in request.files:
        logger.warning("No file part in the request")
        return jsonify({'success': False, 'message': 'No file part in the request'}), 400

    file = request.files['file']

    # Check if file is empty
    if file.filename == '':
        logger.warning("No file selected")
        return jsonify({'success': False, 'message': 'No file selected'}), 400

    # Check file type
    if not allowed_file(file.filename):
        logger.warning(f"Invalid file type: {file.filename}")
        return jsonify({'success': False, 'message': 'File type not allowed. Please upload an Excel file (.xlsx or .xls)'}), 400

    # Save and process the file
    temp_filepath = None
    processed_filepath = None

    try:
        # Get original filename and extension
        original_filename = secure_filename(file.filename)
        file_extension = Path(original_filename).suffix
        file_basename = Path(original_filename).stem

        # Create a unique filename
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        unique_id = str(uuid.uuid4())[:8]  # Use first 8 chars of UUID

        # New filename format: original_name_YYYYMMDD_HHMMSS_uuid.extension
        unique_filename = f"{file_basename}_{timestamp}_{unique_id}{file_extension}"

        # Temporary local filepath for processing
        temp_filepath = os.path.join(
            app.config['UPLOAD_FOLDER'], unique_filename)
        logger.info(f"Saving temporary file to {temp_filepath}")

        # Save file temporarily to process it
        file.save(temp_filepath)

        # Process the Excel file to handle merged cells
        processed_filepath = process_excel_with_merged_cells(temp_filepath)

        # Read the processed Excel file
        logger.info(f"Reading processed Excel file {processed_filepath}")
        df = pd.read_excel(processed_filepath, engine='openpyxl')

        # Check if data exists
        if df.empty:
            # Delete temporary files
            os.remove(temp_filepath)
            if os.path.exists(processed_filepath):
                os.remove(processed_filepath)
            logger.warning("The Excel file contains no data")
            return jsonify({'success': False, 'message': 'The Excel file contains no data'}), 400

        # Convert DataFrame to list of dictionaries and fix numeric keys
        records = df.to_dict('records')
        logger.info(f"Extracted {len(records)} records from Excel file")

        # Process records to convert numeric keys to strings
        processed_records = []
        for record in records:
            processed_record = {}
            for key, value in record.items():
                # Convert any non-string keys to strings
                string_key = str(key) if not isinstance(key, str) else key
                processed_record[string_key] = value
            processed_records.append(processed_record)

        # Store original file in GridFS
        file_id = None
        with open(temp_filepath, 'rb') as f:
            # Store in GridFS with metadata
            file_id = fs.put(
                f.read(),
                filename=unique_filename,
                original_filename=original_filename,
                content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' if file_extension.lower() == '.xlsx' else 'application/vnd.ms-excel',
                upload_date=datetime.datetime.now(),
                month=month,
                year=year
            )
            logger.info(f"Stored file in GridFS with ID {file_id}")

        # Create upload record with reference to GridFS file
        file_size = os.path.getsize(
            temp_filepath) if os.path.exists(temp_filepath) else 0
        upload_info = {
            'filename': unique_filename,
            'original_filename': original_filename,
            'upload_date': datetime.datetime.now(),
            'file_size': file_size,
            'file_id': file_id,
            'record_count': len(processed_records),
            'month': month,
            'year': year
        }

        # Insert upload info to MongoDB
        upload_id = mongo.db.uploads.insert_one(upload_info).inserted_id
        logger.info(f"Created upload record with ID {upload_id}")

        # Overwrite existing records for this month/year - handle duplicate key errors
        try:
            deleted_count = overwrite_roster_for_month_year(
                month, year, processed_records)
            logger.info(
                f"Overwritten {deleted_count} records for {calendar.month_name[month]} {year}")
        except pymongo.errors.BulkWriteError as bwe:
            logger.error(
                f"Bulk write error during record insertion: {bwe.details}")
            return jsonify({
                'success': False,
                'message': f'Error processing file due to duplicate keys. Please try again.'
            }), 500
        except Exception as e:
            logger.error(f"Error overwriting records: {str(e)}")
            return jsonify({
                'success': False,
                'message': f'Error processing file: {str(e)}'
            }), 500

        # Clean up temporary files
        if os.path.exists(temp_filepath):
            os.remove(temp_filepath)
        if os.path.exists(processed_filepath):
            os.remove(processed_filepath)

        overwrote_message = f" (overwrote {deleted_count} existing records)" if deleted_count > 0 else ""

        return jsonify({
            'success': True,
            'message': f'Successfully uploaded {len(processed_records)} records for {calendar.month_name[month]} {year}{overwrote_message}',
            'filename': unique_filename,
            'upload_id': str(upload_id),
            'redirect': url_for('roster', month=month, year=year)
        })

    except Exception as e:
        # Log the full exception with traceback
        logger.error(f"Error during file upload: {str(e)}")
        import traceback
        logger.error(traceback.format_exc())

        # Clean up temporary files
        if temp_filepath and os.path.exists(temp_filepath):
            os.remove(temp_filepath)
        if processed_filepath and os.path.exists(processed_filepath):
            os.remove(processed_filepath)

        return jsonify({'success': False, 'message': f'Error processing file: {str(e)}'}), 500


@app.route('/save-manual-entry', methods=['POST'])
def save_manual_entry():
    """Save a manually entered roster record"""
    try:
        data = request.json

        # Validate required fields
        required_fields = ['month', 'year', 'employee_name',
                           'employee_id', 'department', 'shifts']
        for field in required_fields:
            if field not in data:
                return jsonify({'success': False, 'message': f'Missing required field: {field}'}), 400

        # Validate month and year
        month = data['month']
        year = data['year']

        if month < 1 or month > 12:
            return jsonify({'success': False, 'message': 'Month must be between 1 and 12'}), 400

        if year < 2000 or year > 2100:
            return jsonify({'success': False, 'message': 'Year must be between 2000 and 2100'}), 400

        # Prepare record
        record = {
            'month': month,
            'year': year,
            'employee_name': data['employee_name'],
            'employee_id': data['employee_id'],
            'department': data['department'],
            'upload_id': None  # Manual entry has no upload
        }

        # Add shifts
        shifts = data['shifts']
        if not isinstance(shifts, dict):
            return jsonify({'success': False, 'message': 'Shifts must be a dictionary of day numbers to shift codes'}), 400

        # Get days in month
        days_in_month = calendar.monthrange(year, month)[1]

        # Add each day's shift
        for day in range(1, days_in_month + 1):
            day_str = str(day)
            if day_str in shifts:
                record[day_str] = shifts[day_str]
            else:
                record[day_str] = ''  # Empty for days without shifts

        # Insert the record
        result = mongo.db.users.insert_one(record)

        return jsonify({
            'success': True,
            'message': 'Roster entry created successfully',
            '_id': str(result.inserted_id),
            'redirect': url_for('roster', month=month, year=year)
        })

    except Exception as e:
        logger.error(f"Error saving manual entry: {str(e)}")
        return jsonify({'success': False, 'message': f'Error: {str(e)}'}), 500


@app.route('/user/<user_id>', methods=['GET'])
def get_user(user_id):
    """Get a specific user for editing"""
    try:
        user = mongo.db.users.find_one({'_id': ObjectId(user_id)})

        if not user:
            logger.warning(f"User not found: {user_id}")
            return jsonify({'success': False, 'message': 'User not found'}), 404

        # Convert ObjectId to string for JSON serialization
        user['_id'] = str(user['_id'])
        if 'upload_id' in user and user['upload_id']:
            user['upload_id'] = str(user['upload_id'])

        return jsonify(user)
    except Exception as e:
        logger.error(f"Error retrieving user {user_id}: {str(e)}")
        return jsonify({'success': False, 'message': f'Error retrieving user: {str(e)}'}), 500


@app.route('/user/<user_id>', methods=['PUT'])
def update_user(user_id):
    """Update a specific user"""
    data = request.json

    # Remove _id from the data if it exists
    if '_id' in data:
        del data['_id']

    try:
        result = mongo.db.users.update_one(
            {'_id': ObjectId(user_id)},
            {'$set': data}
        )

        if result.matched_count == 0:
            logger.warning(f"User not found for update: {user_id}")
            return jsonify({'success': False, 'message': 'User not found'}), 404

        logger.info(f"Updated user {user_id}")
        return jsonify({'success': True, 'message': 'User updated successfully'})
    except Exception as e:
        logger.error(f"Error updating user {user_id}: {str(e)}")
        return jsonify({'success': False, 'message': f'Error updating user: {str(e)}'}), 500


@app.route('/user/<user_id>', methods=['DELETE'])
def delete_user(user_id):
    """Delete a specific user"""
    try:
        result = mongo.db.users.delete_one({'_id': ObjectId(user_id)})

        if result.deleted_count == 0:
            logger.warning(f"User not found for deletion: {user_id}")
            return jsonify({'success': False, 'message': 'User not found'}), 404

        logger.info(f"Deleted user {user_id}")
        return jsonify({'success': True, 'message': 'User deleted successfully'})
    except Exception as e:
        logger.error(f"Error deleting user {user_id}: {str(e)}")
        return jsonify({'success': False, 'message': f'Error deleting user: {str(e)}'}), 500


@app.route('/delete-upload/<upload_id>', methods=['POST'])
def delete_upload(upload_id):
    """Delete an upload and all associated records"""
    try:
        # Find the upload
        upload = mongo.db.uploads.find_one({'_id': ObjectId(upload_id)})
        if not upload:
            logger.warning(f"Upload not found for deletion: {upload_id}")
            return jsonify({'success': False, 'message': 'Upload not found'}), 404

        # Delete file from GridFS
        fs.delete(ObjectId(upload['file_id']))
        logger.info(f"Deleted file {upload['file_id']} from GridFS")

        # Delete all users associated with this upload
        result = mongo.db.users.delete_many({'upload_id': upload_id})
        logger.info(
            f"Deleted {result.deleted_count} users associated with upload {upload_id}")

        # Delete the upload record
        mongo.db.uploads.delete_one({'_id': ObjectId(upload_id)})
        logger.info(f"Deleted upload record {upload_id}")

        return jsonify({'success': True, 'message': 'Upload and all associated records deleted successfully'})
    except Exception as e:
        logger.error(f"Error deleting upload {upload_id}: {str(e)}")
        return jsonify({'success': False, 'message': f'Error deleting upload: {str(e)}'}), 500

# ----------------------------
# Error Handlers
# ----------------------------


@app.errorhandler(404)
def not_found(error):
    logger.warning(f"404 error: {request.url}")
    if request.path.startswith('/api/'):
        return jsonify({'error': 'Not found'}), 404
    return render_template('error.html', error="Page not found"), 404


@app.errorhandler(500)
def server_error(error):
    logger.error(f"500 error: {str(error)}")
    if request.path.startswith('/api/'):
        return jsonify({'error': 'Server error'}), 500
    return render_template('error.html', error="Server error occurred"), 500

# ----------------------------
# Main Entry Point
# ----------------------------


if __name__ == '__main__':
    app.run(debug=True)
