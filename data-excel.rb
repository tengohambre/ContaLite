require 'rubygems'
require 'rubyXL'
require 'sqlite3'
require 'active_record'

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

def last_mov
	last = Movimientos.find(:last)
	puts last.inspect
	#puts "El ultimo movimiento fue ",last.idMovimiento,"  ",last.idReferencia,"  ",last.valor
end

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
	 			#puts cell,"its date" ,"\n\n\n"


	 		when 2
	 			@code=cell
	 			#puts cell,"its code","\n\n\n"


	 		when 3
	 			#puts cell,"its shit","\n\n\n"


	 		when 4
	 			if(cell!=nil)
	 				@value=cell
	 				@type=0
	 			end
	 			#puts cell,"its 'egreso'","\n\n\n"


	 		when 5
	 			if(cell!=nil)
	 				@value=cell
	 				@type=1
	 			end
	 			#puts cell,"its ingreso","\n\n\n"


	 		when 6
	 			#puts cell,"its $$$","\n\n\n"
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