require 'rubyXL' /*for processing Excel file*/
require 'rubygems'
require 'rtf'   /*makes rtf document*/
require 'barby' /*makes barcodes*/
require 'barby/barcode/code_128'
require 'barby/outputter/png_outputter'

include RTF

xls = RubyXL::Parser.parse(ARGV[0])  /*accepts the source Excel file*/
data = xls.worksheets[0]

document = Document.new(Font.new(Font::ROMAN, 'Arial'))

i = 1
lines = 0
while defined?(data.sheet_data[i][4].value)  /*breakdown of Excel file*/
	copies = data.sheet_data[i][4].value  /*number of inventory signs varies*/
		
	barcode = Barby::Code128B.new(data.sheet_data[i][3].value) /*converts data to code 128 barcode */
	blob = Barby::PngOutputter.new(barcode)
	blob.height = 50
	blob.xdim = 2
	File.open('barcode.png', 'wb') {|f| f.write blob.to_png}  /*barcode is png file*/
	
	copies.times do  /*builds this label and adds it to the rtf file*/ 
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
		
		lines += 1
		puts lines
		if (lines % 5 == 0)
			document.page_break
		end
	end
	
	i += 1
end

File.open('inventorylabels.rtf', 'w+') {|file| file.write(document.to_rtf)} /*write to the rtf file*/

ObjectSpace.each_object(IO) {|f| f.close unless f.closed? } /*close any and all instances of the png so it can be deleted*/
File.delete('barcode.png')
