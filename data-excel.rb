
require 'green_shoes'
require 'rubygems'
require 'rubyXL'
require 'sqlite3'
require 'active_record'


Shoes.app :title => "ContaLite", :width => 500, :height => 500, :resizable => false do
background "#CAE1D6"

@f1=flow :width => 175, :height => 500 do
	stack :width => 175 do
		para " "
		para "Codigo: ",:align => 'right'
		para "Valor: ",:align => 'right'
		para "",:size => "xx-small"
		para "Fecha(AAAA/MM/DD): ",:align => 'right',:size => "medium"
	end
end

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

	@s=stack :width => 200  do
		para ""

		a = edit_line "Codigo", width: 200
		b = edit_line "Valor", width: 200
		c = edit_line "Fecha (AAAA/MM/DD)", :width => 200
	
		button "Nuevo Movimiento" do
				 	transaccion = Movimientos.new
					transaccion.idMovimiento=nil
					transaccion.idReferencia=a.text()
					transaccion.valor=b.text();
					transaccion.fecha=c.text();
					transaccion.save
					puts "Se Agrego ",transaccion.inspect
		end
	
		para "  "
		
		button "Load from EXCEL" do
			lectura=1
			if(lectura==1)
				filename=ask_open_file
					#workbook = RubyXL::Parser.parse("Book1.xlsx")
					workbook = RubyXL::Parser.parse(filename)
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

		button "ADD Code/Desc" do
			window :title => "New Code adder" do
				background "#DFA"
				flow :width => 175, :height => 500 do
					stack :width => 175 do
						para " ", :align => "left"
						para "Codigo: ", :align => "right"
						para " ", :align => "left", :size => "xx-small"
						para "Descr: ", :align => "right"

					end
				end
				stack :width => 200 do
					para " "
					@codigo = edit_line "codigo", :width => 200
					@desc = edit_line "Descripcion", :width => 200

					button "Agregar Codigo" do
						
						nuCode= Referencias.new
						nuCode.id=@codigo.text
						nuCode.desc=@desc.text
						nuCode.save
						puts "Se agrego ",nuCode.inspect
					end
				end

			end
		end

		 button ('Close') {exit}

	
	end


end
