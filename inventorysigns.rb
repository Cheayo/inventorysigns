require 'rubyXL'
require 'rubygems'
require 'rtf'
require 'barby'
require 'barby/barcode/code_128'
require 'barby/outputter/png_outputter'

include RTF

xls = RubyXL::Parser.parse(ARGV[0])
data = xls.worksheets[0]

document = Document.new(Font.new(Font::ROMAN, 'Arial'))

i = 1;
while defined?(data.sheet_data[i][4].value)
	copies = data.sheet_data[i][4].value
	
	barcode = Barby::Code128B.new(data.sheet_data[i][3].value)
	blob = Barby::PngOutputter.new(barcode)
	blob.height = 50
	blob.xdim = 2
	File.open('barcode.png', 'wb') {|f| f.write blob.to_png}
	
	copies.times do 
		document.paragraph do |p|
			p << "QMI\t\tMARKER"
			p.line_break
			p << data.sheet_data[i][2].value
			p.line_break
			p << data.sheet_data[i][1].value
		end
		
		image = document.image('barcode.png')
		document.paragraph do |p|
			p.line_break
		end
	end
	
	i += 1
end

File.open('inventorylabels.rtf', 'w+') {|file| file.write(document.to_rtf)}
