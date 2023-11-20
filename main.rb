require 'googleauth'
require 'google/apis/drive_v3'
require 'google/apis/sheets_v4'
require 'pry'
require 'mini_magick'

ORIGINAL_FOLDER_ID='1U2Xmb078-icd-IH4pCWQc-sqW2jRMT9MYwDIJ681A9IfwS1y64QB8ArLXYHnnF0ArXhl7XwR'
BACKUP_FOLDER_ID='1W3TY-qJ31THYQdxEi2Q5tjQ-wUCRtemE'
CONVERTED_IDS_FILE_NAME = 'converted.db'
SPREADSHEET_ID='1QbFMeV09e-rcJmoBax4445v3OGZ-38s6aju8zaM30-8'
IMPORT_RANGE = "Form Import!F1:K50"

Drive = ::Google::Apis::DriveV3
drive_service = Drive::DriveService.new
Sheets = Google::Apis::SheetsV4 
sheets_service = Sheets::SheetsService.new

scope = 'https://www.googleapis.com/auth/drive'

authorizer = Google::Auth::ServiceAccountCredentials.from_env(scope: scope)
drive_service.authorization = authorizer
sheets_service.authorization = authorizer

# Confiuration finished


main_sheet = sheets_service.get_spreadsheet(SPREADSHEET_ID, ranges: IMPORT_RANGE, include_grid_data: true)


converted_ids_file = File.open(CONVERTED_IDS_FILE_NAME, 'a+')
converted_ids = converted_ids_file.readlines.map(&:chomp)

form_data = main_sheet.sheets.first.data.first

form_data.row_data.each_with_index do |tab, i|
  p "Processing row: #{i}"
  url = tab.values[0]
  next if url.effective_value.nil?

  str = url.effective_value.string_value
  path = str

  #Select first photo when uploaded many
  if path.include?(',')
    path = path.split(',').first
  end

  file_id = str[/id=\S*/]
  next if path.nil?

  file_id.gsub!(/("|,)/, '')
  file_id.gsub!('id=', '')
  next if converted_ids.include?(file_id)

  image = drive_service.get_file(file_id, fields: 'id,mime_type,web_content_link,name')

  if image.mime_type == 'image/heic' or image.mime_type == 'image/heif'
    #Backup original file
    data = {name: "Original - #{image.name}", parents: [BACKUP_FOLDER_ID]}
    newfile = drive_service.copy_file(file_id, file_object=data)
    p "Backup ID: #{newfile.id}"

    image_file = MiniMagick::Image.open(image.web_content_link)
    image_file.format 'jpg'

    r = drive_service.create_file(file_object={name: image.name.gsub(/(.HEIC|.heic)/, '.jpg'), parents: [ORIGINAL_FOLDER_ID]}, upload_source: image_file.path)
    if r.class == Google::Apis::DriveV3::File
      p 'Converted!'
      
      converted_ids_file.puts file_id

      converted_file_id = r.id
      converted_file = drive_service.get_file(converted_file_id, fields: 'id,web_content_link')
      update_values_request = Google::Apis::SheetsV4::BatchUpdateValuesRequest.new(
        value_input_option: 'RAW',
        data: [
          Google::Apis::SheetsV4::ValueRange.new(range: "Form Import!K#{i+1}", values: [[converted_file.web_content_link]])
        ]
      )
      sheets_service.batch_update_values(SPREADSHEET_ID, update_values_request)
    end
  else
    # push id of jpgs to make things faster
    converted_ids_file.puts file_id
  end
end


