require 'rubyXL'
require 'rubygems'
require 'rtf'

include RTF

xls = RubyXL::Parser.parse(ARGV[0]) #'c:/users/lessa/desktop/inventory labels data set CLEAN.xlsx'
data = xls.worksheets[0]

document = Document.new(Font.new(Font::ROMAN, 'Arial'))

styles = {}
styles['BAR_CODE'] = CharacterStyle.new
styles['BAR_CODE'].font = Font.new(Font::NIL, 'Wingdings')

i = 1;

while defined?(data.sheet_data[i][4].value)
	copies = data.sheet_data[i][4].value
		
	copies.times do 
		document.paragraph do |p|
			p << "QMI\t\tMARKER"
			p.line_break
			p << data.sheet_data[i][2].value
			p.line_break
			p << data.sheet_data[i][1].value
			p.line_break
			p.apply(styles['BAR_CODE']) do |bc|
				bc << "\t" + data.sheet_data[i][3].value
			end
		end
	end
	
	i += 1
end

File.open('inventorylabels.rtf', 'w+') {|file| file.write(document.to_rtf)}
