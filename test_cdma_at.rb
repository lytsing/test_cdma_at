#! /usr/bin/env ruby
#
# test_cdma_at.rb
# @date 2009.5.16
#

require 'mscomm.rb'
require 'excel.rb'

def exec_cmd(comm, cmd)
  comm.write("#{cmd}\r\n")
  sleep(0.2)

  begin
    result = comm.read
    result = "Command not support\n" if result.include?("ERROR\n")
  rescue
    result = "Writing serial port error\n"
  end

  return result
end

atcmd = %w{
AT
AT+PSS?
AT+CPIN?
AT+CPINC?
AT+ARSI=1
AT+GMI
AT+GESN?
AT+CIMI?
AT+CSQ?
AT+CREG=2
AT+CREG?
AT+VMCC?;+VMNC?
AT+CPBS=?
AT+CPBS?
AT+CPBS="ME"
AT+CPBS?
AT+CPBW=3,13544049382,"violet",0
AT+CPBR=3
AT+CPBw=3
AT+CPMS?
AT+CPMS=?
AT+CMGF=?
AT+CMGF?
AT+CMGF=1
AT+CMGS=13544049382,"Hello!"
AT+CMGR=3
AT+CMGD=3
AT+CMGW=,13544049382,"Hi!"
AT+CMGD=4
AT+CMGW=4,13544049382,"Hi!"
AT+CMGR=4
AT+CNUM?
AT+ISF?
AT+VSPST?
AT+CLIP=1
AT+SPEAKER=1
AT+VGT=?
AT+VGT?
AT+VGT=6
AT+VGR=?
AT+VGR?
AT+VGR=6
AT+CLCC?
AT+CPOF
AT+CPON
}

comm  = MSCOMM.new(6)
excel = Excel.new("d:\\Book1.xls")

i = 1

atcmd.each {|at|
  excel.setvalue("a#{i}", at)
  excel.excelobj.Range("b#{i}").Value = exec_cmd(comm, at)
  i += 1
}

comm.close
excel.save
excel.close

