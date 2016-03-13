require 'spreadsheet'
require 'mysql2'

class Record
	attr_accessor :No
	attr_accessor :certificateNo
	attr_accessor :recordTitle
	attr_reader :pageNum
	attr_reader :comment
	attr_accessor :numOfPages
	def set_pageNum(str)
		@pageNum = str
		@numOfPages = parsePageNum(@pageNum)
	end
	def processDatetime(datetime)
			resultMon = datetime.month <10?"0#{datetime.mon}":datetime.month
			resultDay = datetime.day <10?"0#{datetime.day}":datetime.day
			resultHour = datetime.hour <10?"0#{datetime.hour}":datetime.hour 
			resultMin = datetime.min <10?"0#{datetime.min}":datetime.min
			resultSec = datetime.sec <10?"0#{datetime.sec}":datetime.sec
		"'#{datetime.year}-#{resultDay}-#{resultDay} #{resultHour}-#{resultMin}-#{resultSec}'"
	end
	def set_comment(sth)
		if sth.is_a?String
			@comment = sth
		else
			@comment = processDatetime(sth)
		end

	end
	def initialize 
		@No = 0.to_f
		@certificateNo = ""
		@recordTitle = ""
		@pageNum = ""
		@comment = nil
		@numOfPages = parsePageNum(@pageNum)
	end
	def parsePageNum(str)
		if str!=""
			temp = (str[-3..-1].to_i)-(str[0..2].to_i)
			temp+1
			else
				0
		end
	end
	def to_s
		"[室编卷号:#{@No},结婚证字号:#{@certificateNo},文件题目:#{@recordTitle},页号:#{@pageNum},备注:#{@comment},页数:#{@numOfPages}]"
	end
	def to_normal_s
		"(#{@No},'#{@certificateNo}','#{@recordTitle}','#{@pageNum}',#{@comment},#{@numOfPages});"
	end
end

class TransferToDB
	@book = nil
	@path = ""
	@client = nil
	def initialize 
		#temparily path is "/Users/bellchet58/Downloads/婚登档案/2007年婚姻登记档案.xls"
		#which can be modified in the future
		@path = "/Users/bellchet58/Downloads/openofficeVer.xls" 
		Spreadsheet.client_encoding = 'UTF-8'
		@book = Spreadsheet.open(@path) 
		initializeDB
	end
	def formatRecordTitle(str)
		if str.rindex('；')
			str[str.rindex('；')] = "："
		end
		str.split("：")[1][0..-2]+" "+str.split("：")[2]
	end
	def initializeDB
		host = "localhost"
		username = "root"
		database = "archiver"
		@client = Mysql2::Client.new :host=>host, :username=>username, :database=>database
	end
	def readyForRecords(tableName)
		result = @client.query "show tables;",:as => :array
		arr = []
		result.each do |row|
			arr << row[0]
		end
		if arr.any? { |name|  name== tableName }
			@client.query "drop table #{tableName};"
		end
		@client.query "CREATE TABLE #{tableName} ( No float, certificate varchar(100), recordTitle varchar(100), pageNum char(7), comment datetime, numOfPages int)charset=utf8;"
	end
	def action
		#do something in the loop of boundsheets
		  #do something in a single boundsheet
		    #若 sheet*.row(*)[0] 为“结婚登记档案目录时，则游标向下 3 行
		    #读出编号、证字号、页数、备注、新增？ 姓名	
		readyForRecords("marriage_records")
		@book.worksheets.each do |sheet|
			# 删除无用行
			# 或是跳过
			
			# 其实可以用Object.clone复制新的纪录
			tempNo = 0.to_f
			tempCertificateNo = ""
			tempRecordTitle = ""
			tempPageNum = ""
			tempComment = nil


			nextLineIsUseful = true
			# 暂时将纪录存为数组
			records = []
			
			for i in 0..sheet.last_row_index
				record = Record.new
				if sheet.row(i).at(0) == '结婚登记档案目录 ' 
					nextLineIsUseful = false
					next
				elsif !nextLineIsUseful
					nextLineIsUseful = true
					next
				elsif sheet.row(i).at(0) == '室编卷号'
					next
				elsif sheet.row(i).at(3) == nil
					next
				end
				# puts sheet.row(i)
				if sheet.row(i).at(0)!=nil
					puts "#{i}为纪录第一行"
					if sheet.row(i).at(0).is_a?String
						tempNo = sheet.row(i).at(0).to_f
					else
						tempNo = sheet.row(i).at(0)
					end		
					tempCertificateNo = sheet.row(i).at(1)
					tempRecordTitle = sheet.row(i).at(2)+sheet.row(i).at(3)
					tempPageNum = sheet.row(i).at(4)
					if sheet.row(i).at(5).is_a?Float
						tempComment = sheet.row(i).datetime(5)
						else
							tempComment = sheet.row(i).at(5)
					end
					
					elsif sheet.row(i).at(0)==nil		
						puts "#{i}为纪录第二行"
						record.No = tempNo
						record.certificateNo = tempCertificateNo
						record.recordTitle = formatRecordTitle(tempRecordTitle+sheet.row(i).at(2)+sheet.row(i).at(3))
						record.set_pageNum tempPageNum

						if tempComment==nil || tempComment==""
							record.set_comment "null"
						elsif 
							record.set_comment tempComment
						end
						puts "insert into marriage_records values "+record.to_normal_s
						records << record
						@client.query "insert into marriage_records values "+record.to_normal_s
				end
			end
			puts records
		end
		
	end
end
TransferToDB.new.action