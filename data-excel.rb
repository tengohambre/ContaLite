
require 'green_shoes'
require 'rubygems'
require 'rubyXL'
require 'sqlite3'
require 'active_record'


Shoes.app :title => "ContaLite", :width => 500, :height => 500, :resizable => false do

	ActiveRecord::Base.establish_connection(
		:adapter => "sqlite3",
		:database => "sample.db"
	)

	class Movimientos < ActiveRecord::Base
		belongs_to :referencias
	end

	class Referencias < ActiveRecord::Base
		has_many :movimientos
	end

	button "Load from EXCEL" do
		lectura=1
		if(lectura==1)
				workbook = RubyXL::Parser.parse("Book1.xlsx")
				array=workbook.worksheets[0].extract_data

				array.each do |row|
	 				i=1;
	 				@date=""
	 				@code=""
	 				@value=""
	 				@type=""

		 			row.each do |cell|

						case i
					 		when 1
	 							@date=cell
					 		when 2
	 							@code=cell
				 			when 3
					 		when 4
					 			if(cell!=nil)
	 								@value=cell
	 								@type=0
	 							end
				 			when 5
	 							if(cell!=nil)
	 								@value=cell
	 								@type=1
		 						end
					 		when 6
					 	end

	 					i+=1
	 				end

		 			if(@type==0)
		 				@value=@value*-1
	 				end

	 				print @date," ",@code," ",@value," ",@type,"  \n"
				 	transaccion = Movimientos.new
					transaccion.idMovimiento=nil
					transaccion.idReferencia=@code
					transaccion.valor=@value
					transaccion.fecha=@date
					transaccion.save
				end

		end
	end
end
