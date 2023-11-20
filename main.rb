require 'googleauth'
require 'google/apis/drive_v3'
require 'google/apis/sheets_v4'
require 'pry'
require 'mini_magick'

ORIGINAL_FOLDER_ID='1U2Xmb078-icd-IH4pCWQc-sqW2jRMT9MYwDIJ681A9IfwS1y64QB8ArLXYHnnF0ArXhl7XwR'
BACKUP_FOLDER_ID='1W3TY-qJ31THYQdxEi2Q5tjQ-wUCRtemE'
CONVERTED_IDS_FILE = 'converted.db'


Drive = ::Google::Apis::DriveV3
drive_service = Drive::DriveService.new
Sheets = Google::Apis::SheetsV4 
sheets_service = Sheets::SheetsService.new


scope = 'https://www.googleapis.com/auth/drive'

authorizer = Google::Auth::ServiceAccountCredentials.from_env(scope: scope)
drive_service.authorization = authorizer
sheets_service.authorization = authorizer

list_heif_files = drive_service.list_files(q: "mimeType = 'image/heif'")

# Get spreadsheet
all_sheets = drive_service.list_files(q: 'mimeType = "application/vnd.google-apps.spreadsheet"')
# and ID
sheet_id = all_sheets.files.last.id
range = "Form Import!F1:K100"
main_sheet = sheets_service.get_spreadsheet(sheet_id, ranges: range, include_grid_data: true)

# str = main_sheet.sheets.first.data.first.row_data[1].values[4].effective_value.string_value 

converted_ids_file = File.open(CONVERTED_IDS_FILE, 'a+')
converted_ids = converted_ids_file.readlines.map(&:chomp)

csv_data = main_sheet.sheets.first.data.first

csv_data.row_data.each_with_index do |tab, i|
  p "Processing row: #{i}"
  url = tab.values[0]

  next if url.effective_value.nil?

  str = url.effective_value.string_value

  #path = str[/https\S*"/]
  path = str
  if path.include?(',')
    path = path.split(',').first
  end

  file_id = str[/id=\S*/]
  next if path.nil?


  #path.gsub!('"', '')


  file_id.gsub!(/("|,)/, '')
  file_id.gsub!('id=', '')

  next if converted_ids.include?(file_id)

  image = drive_service.get_file(file_id, fields: 'id,mime_type,web_content_link,name')

  if image.mime_type == 'image/heic' or image.mime_type == 'image/heif'
    #Backup original file
    data = {name: "Original - #{image.name}", parents: [BACKUP_FOLDER_ID]}
    newfile = drive_service.copy_file(file_id, file_object=data)

    p "Backup ID: #{newfile.id}"
    # binding.pry
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
      sheets_service.batch_update_values(sheet_id, update_values_request)
    end
  else
    # push id of jpgs to make things faster
    converted_ids_file.puts file_id
  end
end


