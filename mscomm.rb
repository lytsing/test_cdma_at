#! /usr/bin/env ruby
#
# mscomm.rb
# @date 2009.5.16
#

require 'win32ole'

class MSCOMM
  def initialize(port)
    @serial = WIN32OLE.new("MSCOMMLib.MSComm")

    @serial.CommPort = port
    @serial.Settings = "115200,N,8,1"
    @serial.InputLen = 0
    @serial.PortOpen = true
  end

  def write(str)
    @serial.Output = str
  end

  def read
    str = @serial.Input
    str
  end

  def close
    @serial.PortOpen = false
  end

  def serial
    @serial
  end
end

