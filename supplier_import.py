#coding:utf-8

from openerp.osv import osv,fields
from openerp.tools.translate import _
import xlrd,base64
import time

class supplier_import(osv.osv):
	_name="supplier.import"
	_columns={
                "xls":fields.binary('Excel File'),
                }

	def btn_import(self,cr,uid,ids,context=None):
			for wiz in self.browse(cr,uid,ids):
					if not  wiz.xls:
							continue
					excel = xlrd.open_workbook(file_contents=base64.decodestring(wiz.xls))
					sheets = excel.sheets()
					for sh in sheets:
							for row in range(2,sh.nrows):
									print sh.cell(row,0).value
									if sh.cell(row,0).value=="Y":
											is_company = True
									else:
											is_company = False
									if sh.cell(row,1).value=="Y":
											is_department = True
									else:
											is_department = False
									if not sh.cell(row,2).value:
											raise osv.except_osv(_('Error!'),_('feild name is required !'))
									else:
											name=sh.cell(row,2).value
											
									#address
									addresses = sh.cell(row,3).value.split(',')
									if len(addresses) <5:
											raise osv.except_osv(_('Error!'),_('address is not compelted!'))
									else:
											countries = self.pool.get('res.country').search(cr,uid,[('name','=',addresses[0])],context=context)
											if len(countries):
													country = countries[0]
											else:
													raise osv.except_osv(_('Error!'),_('The country is not existed!'))

											states =  self.pool.get('res.country.state').search(cr,uid,[('name','=',addresses[1])],context=context)
											if len(states):
													state = states[0]
											else:
													raise osv.except_osv(_('Error!'),_('The state is not existed!'))

											cities = self.pool.get('rainsoft.city').search(cr,uid,[('name','=',addresses[2])],context=context)
											if len(cities):
													city = cities[0]
											else:
													raise osv.except_osv(_('Error!'),_('The city is not existed!'))
													
											districts = self.pool.get('rainsoft.district').search(cr,uid,[('name','=',addresses[3])],context=context)

											if len(districts):
													district= districts[0]
											else:
													raise osv.except_osv(_('Error!'),_('The city is not existed!'))

													
											street = addresses[4]
									tel = sh.cell(row,4).value
									phone = sh.cell(row,5).value
									email = sh.cell(row,6).value
									QQ = str(sh.cell(row,7).value).split('.')[0]
									comment = sh.cell(row,8).value
									no = str(sh.cell(row,9).value).split('.')[0]
									date_start = sh.cell(row,10).value
									date_end = sh.cell(row,11).value

									line={
													"is_company":is_company,
													"name":name,
													"is_internal":is_department,
													"country_id":country,
													"state_id":state,
													"city":city,
													"district":district,
													"street":street,
													"phone":tel,
													"mobile":phone,
													"email":email,
													"QQ":QQ,
													"comment":comment,
													"ref":no,
													"date":date_start,
													"contract_end_date":date_end,
													"supplier":True,
													"active":True,
													"customer":False,
													}
									print line
									self.pool.get('res.partner').create(cr,uid,line,context=context)




			

