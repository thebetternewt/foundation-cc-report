require "fileutils"
require "sqlite3"
require 'csv'

# Return string current date and time.
def current_datetime
  d = DateTime.now
  d.strftime("%Y-%m-%d_%I_%M_%S")
end

def add_new_gift(gifts_array, new_gift)
  gifts_array << new_gift
end

def fields_for_sql(num_fields)
  '(' + "?,\s" * (num_fields - 1) + '?' + ')'
end

def insert_gift_into_csv(gift, csv)
        csv << [
          gift['settle_date'],
          gift['last_name'],
          gift['first_name'],
          gift['c_last_name'],
          gift['c_first_name'],
          gift['banner_id'],
          # gift['donor_id'],
          # gift['total_gift_amount'],
          gift['pledge_number'],
          gift['designation_amount'],
          gift['desg_code'],
          gift['other_designation'],
          # gift['gift_amount'],
          # gift['gift_designation'],
          # gift['gift_amount2'],
          # gift['gift_designation2'],
          gift['gift_description'],
          gift['solicitation_code'],
          # gift['c_solicitation_code'],
          # gift['trans_id'],
          # gift['transaction_id'],
          gift['anonymous'],
          gift['gift_matching'],
          gift['tribute_type'],
          gift['tribute_occasion'],
          gift['tribute_notification_name'],
          gift['tribute_notification_address'],
          gift['tribute_comments'],
          gift['mem_in_honor'],
          gift['next_of_kin'],
          gift['comments'],
          gift['card_description'],
          gift['tran_type'],
          gift['user_id'],
          gift['batch_num']
        ]
end

# Create directories for reports
Dir.mkdir('reports') unless File.exists?('reports')
timestamp = current_datetime

# Create Dtabase
db = SQLite3::Database.new ":memory:"
db.results_as_hash = true

# Create tables
db.execute <<-SQL
  create table designations (
    gift_id int primary key,
    last_name varchar(30),
    first_name varchar(30),
    banner_id varchar(10),
    date_stamp varchar(30),
    transaction_id varchar(10),
    designation_amount varchar(10),
    desg_code varchar(30)
  );
SQL

db.execute <<-SQL
  create table gift_info (
    id int primary key,
    last_name varchar(30),
    first_name varchar(30),
    address_1 varchar(100),
    address_2 varchar(100),
    city varchar(100),
    state varchar(10),
    zip varchar(20),
    phone_type varchar(10),
    area varchar(5),
    phone_number varchar(15),
    email varchar(50),
    anonymous varchar(10),
    other_designation varchar(30),
    solicitation_code varchar(10),
    gift_matching varchar(30),
    tribute_type varchar(10),
    tribute_full_name varchar(10),
    tribute_occasion varchar(30),
    tribute_notification_name varchar(30),
    tribute_notification_address varchar(30),
    tribute_comments varchar(100),
    date_submitted varchar(30),
    trans_number varchar(100)
  );
SQL

db.execute <<-SQL
  create table converge_payments (
    transaction_id varchar(100) primary key,
    settle_date varchar(30),
    user_id varchar(20),
    card_description varchar(10),
    gift_description varchar(30),
    first_name varchar(30),
    last_name varchar(30),
    donor_id varchar(10),
    address_1 varchar(100),
    address_2 varchar(100),
    city varchar(100),
    state varchar(10),
    zip varchar(20),
    phone_number varchar(15),
    email varchar(50),
    gift_designation varchar(10),
    gift_designation2 varchar(10),
    comments varchar(200),
    mem_in_honor varchar(100),
    next_of_kin varchar(100),
    pledge_number varchar(6),
    total_gift_amount varchar(20),
    gift_amount varchar(20),
    gift_amount2 varchar(20),
    solicitation_code varchar(20),
    tran_type varchar(10),
    batch_num varchar(10)
  );
SQL

# Populate iModules Designations table.
CSV.foreach('imod_desg.csv', headers: true) do |row|
  db.execute "INSERT INTO designations (
    gift_id,
    last_name,
    first_name,
    banner_id,
    date_stamp,
    transaction_id,
    designation_amount,
    desg_code) VALUES #{fields_for_sql(8)}",
    [ row['ID'],
      row['Last Name'],
      row['First Name'],
      row['Banner_ID'],
      row['Date Stamp'],
      row['Transaction ID'],
      row['Designation Amount'],
      row['ADBDESG_DESG']
    ]
end

# Populate iModules Export table.
CSV.foreach('imod_export.csv', headers: true) do |row|
  db.execute "INSERT INTO gift_info VALUES #{fields_for_sql(24)}",
    [ row['Transaction ID'],
      row['Last Name'],
      row['First Name'],
      row['Address_1'],
      row['Address_2'],
      row['City'],
      row['State'],
      row['Zip'],
      row['imod_phone_type'],
      row['Area'],
      row['Phone_Number'],
      row['Primary E-mail'],
      row['MAG12 - Is Anonymous'],
      row['MAG12 - OtherDesignation'],
      row['Giving - Solicitation Type'],
      row['Make a Gift - MAG12 - Gift Matching'],
      row['MAG12 - TributeType'],
      row['MAG12 - TributeFullName'],
      row['MAG12 - TributeOccasion'],
      row['MAG12 - TributeNotificationName'],
      row['MAG12 - TributeNotificationAddress'],
      row['MAG12 - TributeComments'],
      row['date_submitted'],
      row['Customer Trans Number'] ]
end

# Join iModules designations and export tables with needed fields.
imod_query = "SELECT
      gift_info.first_name,
      gift_info.last_name,
      banner_id,
      address_1,
      address_2,
      city,
      state,
      zip,
      area,
      phone_type,
      phone_number,
      email,
      designation_amount,
      desg_code,
      other_designation,
      solicitation_code,
      id as trans_id,
      trans_number,
      anonymous,
      gift_matching,
      tribute_type,
      tribute_occasion,
      tribute_notification_name,
      tribute_notification_address,
      tribute_comments
    FROM gift_info
    LEFT OUTER JOIN designations ON trans_id = transaction_id
    WHERE trans_id NOT NULL"

imod_records = db.execute imod_query

# Write iModules query results to new file.
CSV.open('imod_report.csv', 'w') do |csv|
  # Headers
  csv << [
    'Last Name',
    'First Name',
    'Banner_ID',
    'Designation Amount',
    'Designation Code',
    'Other Designation',
    'Solicitation Code',
    'Transaction ID',
    'Transaction Number',
    'Anonymous',
    'Gift Matching',
    'Tribute Type',
    'Tribute Occasion',
    'Tribute Notification Name',
    'Tribute Notification Address',
    'Tribute Comments'
  ]

  # Data
  imod_records.each do |record|
    csv << [ record['last_name'],
             record['first_name'],
             record['banner_id'],
             record['designation_amount'],
             record['desg_code'],
             record['other_designation'],
             record['solicitation_code'],
             record['trans_id'],
             record['trans_number'],
             record['anonymous'],
             record['gift_matching'],
             record['tribute_type'],
             record['tribute_occasion'],
             record['tribute_notification_name'],
             record['tribute_notification_address'],
             record['tribute_comments']
            ]
  end
end
puts
puts "CSV file 'imod_report.csv' created!"

# Read in Converge Batch Report overall totals as variable.
converge_report_overall_totals = CSV.read('converge_report.csv').last

# Read in Converge Batch Report and strip leading whitespace from values.
converge_report = CSV.read('converge_report.csv',
  skip_blanks: true,
  skip_lines: '(Detail report*)|(Created on*)|(Overall Totals*)')



converge_report.each do |row|
  row.each { |value| value.to_s.lstrip! }
end

# Write clean Converge Batch Report data to new file 'new_converge_report.csv'.
CSV.open('new_converge_report.csv', 'w') do |csv|
  converge_report.each { |row| csv << row }
end

puts
puts "CSV file 'new_converge_report.csv' created!"


# Populate converge_payments table from new file.
CSV.foreach('new_converge_report.csv', headers: true) do |row|

  db.execute "INSERT INTO converge_payments VALUES #{fields_for_sql(27)}",
    [ row['Transaction'],
      row['Settle Date'],
      row['User ID'],
      row['Card Description'],
      row['Description'],
      row['First Name'],
      row['Last Name'],
      row['Donor ID'],
      row['Address1'],
      row['Address2'],
      row['City'],
      row['State/Province'],
      row['Postal code'],
      row['Phone'],
      row['Email Address'],
      row['Gift Designation'],
      row['Gift Designation 2'],
      row['Comments'],
      row['Memorial In Honor Of'],
      row['Next of Kin'],
      row['Pledge Number'],
      row['Amount'],
      row['Gift Amount'],
      row['Gift Amount 2'],
      row['Solicitation Code'],
      row['Tran Type'],
      row['Batch Number']
    ]

end

# Query for for final giving report.
gifts = db.execute \
  "SELECT * FROM (
    SELECT
        *, last_name as c_last_name,
        first_name as c_first_name,
        address_1 as c_address_1,
        address_2 as c_address_2,
        city as c_city,
        state as c_state,
        zip as c_zip,
        phone_number as c_phone_number,
        email as c_email,
        solicitation_code as c_solicitation_code
      FROM converge_payments
  ) LEFT OUTER JOIN (#{imod_query}) ON trans_number = transaction_id
    ORDER BY user_id DESC"

# Write final report data to file.
filename = "#{timestamp}_gift_admin.csv"

CSV.open("reports/#{filename}", 'w') do |csv|
  # Headers
  csv << [ 'Settle Date',
           'Last Name',
           'First Name',
           'C_Last Name',
           'C_First Name',
           'Banner_ID',
           # 'C_Donor ID',
           # 'C_Total GIft Amount',
           'Pledge Number',
           'Designation Amount',
           'Designation Code',
           'Other Designation',
           # 'C_Gift Amount 1',
           # 'C_Gift Designation 1',
           # 'C_Gift Amount 2',
           # 'C_Gift Designation 2',
           'C_Description',
           'Solicitation Code',
           # 'C_Solicitation Code',
           # 'Transaction ID',
           # 'Transaction Number',
           'Anonymous',
           'Gift Matching',
           'Tribute Type',
           'Tribute Occasion',
           'Tribute Notification Name',
           'Tribute Notification Address',
           'Tribute Comments',
           'C_Mem/In Honor of',
           'C_Next of Kin',
           'C_Comments',
           'C_Card Description',
           'Tran Type',
           'C_User ID',
           'C_Batch #'
          ]

  # -----------
  # Insert Data
  # -----------

  # Sort gifts by 'settle_date'
  gifts.sort_by! { |gift| gift['settle_date']}

  gifts.each do |gift|

    # Update card description.
    case gift['card_description']
      when 'VISA'
        gift['user_id'] == 'Webpage' ? gift['card_description'] = 'WM' : gift['card_description'] = 'MC'
      when 'AMEX'
        gift['user_id'] == 'Webpage' ? gift['card_description'] = 'WX' : gift['card_description'] = 'AX'
      when 'DISC'
        gift['user_id'] == 'Webpage' ? gift['card_description'] = 'WD' : gift['card_description'] = 'DS'
    end

    # iModules Tribute Type codes
    case gift['tribute_type']
      when 'In Memory'
        gift['tribute_type'] = 'MEMR'
      when 'In Honor'
        gift['tribute_type'] = 'HONR'
    end 
    
    # Anonymous codes = 'ANON'
    gift['anonymous'] == 'True' ? gift['anonymous'] = 'ANON' : gift['anonymous'] = '' 

    # Merge gift/donor info from iModules and Converge.
    gift['banner_id'] = gift['donor_id'] if gift['banner_id'].nil?
    gift['designation_amount'] ||= ( gift['gift_amount'].empty? ? gift['total_gift_amount'] : gift['gift_amount'] )
    gift['desg_code'] = gift['gift_designation'] if gift['desg_code'].nil?
    gift['solicitation_code'] = gift['c_solicitation_code'] if gift['solicitation_code'].nil?

    # Break apart multiple designations.
    if !gift['gift_amount2'].empty?
      # new_gift = gift.to_a
      new_gift = (gift.to_a).to_h
      new_gift['designation_amount'] = gift['gift_amount2']
      new_gift['desg_code'] = gift['gift_designation2']

      # Insert new_gift into csv.
      insert_gift_into_csv(new_gift, csv)

    end

    # Insert gift into csv.
    insert_gift_into_csv(gift, csv)
  end

  # Insert total data from Converge into csv.
  csv << ['']
  csv << converge_report_overall_totals

  puts
  puts "CSV file '#{filename}' created!"

end

# Write data services report data to file.
dataserv_report = "#{timestamp}_data_serv.csv"

CSV.open("reports/#{dataserv_report}", 'w') do |csv|
  csv << [
    'Settle Date',
    'Donor ID',
    # 'test_name', # use for possible blanks?
    'Last Name',
    'First Name',
    'C_Last Name',
    'C_First Name',
    'Address 1',
    'Address 2',
    'City',
    'State',
    'Zip',
    'Phone Type',
    'Phone',
    'Email',
    'C_Address 1',
    'C_Address 2',
    'C_City',
    'C_State',
    'C_Zip',
    'C_Phone',
    'C_Email'
  ]

  # -----------
  # Insert Data
  # -----------

  gifts.each do |gift|
    csv << [
      gift['settle_date'],
      gift['banner_id'],
      # gift['123'] || gift['last_name'], # use for possible blanks?
      gift['last_name'],
      gift['first_name'],
      gift['c_last_name'],
      gift['c_first_name'],
      gift['address_1'],
      gift['address_2'],
      gift['city'],
      gift['state'],
      gift['zip'],
      gift['phone_type'],
      "#{gift['area']}\s#{gift['phone_number']}",
      gift['email'],
      gift['c_address_1'],
      gift['c_address_2'],
      gift['c_city'],
      gift['c_state'],
      gift['c_zip'],
      gift['c_phone_number'],
      gift['c_email']

    ]

  end

  puts
  puts "CSV file '#{dataserv_report}' created!"

end

puts

