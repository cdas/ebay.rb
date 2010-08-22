#!/usr/bin/env ruby
# Copyright 2010 (c) René Wilhelm <rene dot wilhelm at gmail dot com>
# Created: Sat Jul 10 02:39:53 CEST 2010

require 'rubygems'
require 'mechanize'

userid = "cowdaseking"
pass = "legende_31"

agent = Mechanize.new
page = agent.get('https://signin.ebay.de/ws/eBayISAPI.dll?SignIn')

form = agent.page.form('SignInForm')
form.userid = ARGV[0] || userid
form.pass = ARGV[1] || pass
page = form.submit

url = "http://completed.shop.ebay.de/i.html?_nkw=wow+account&LH_Complete=1&_dmd=7&_ipg=200"
page = agent.get(url)

pgns = page.search("//span[@class='page']").inner_html.split.last.to_i

article_numbers = Array.new

(1..pgns).each do |pgn|
  page = agent.get(url + "&_pgn=" + pgn.to_s)
  page.search("//span[@class='v']").each do |article_number|
    article_numbers << article_number.inner_html
  end
  break
end

#
# REGEXEN
#
# Spieldauer:
# Preis:
# Craft:
# Dual:
# LVL:
# Reiten:
# Class Quests Done:
# Conquests:
# Ehre:
# EQ:
# Flugpunkte Done:
# Geld:
# Hero Done:
# Mount:
# Ruf:
# Triumph:
# Auktionstext:

class Auction
  attr_accessor :article_number, :href, :success, :price
  
  def initialize(article_number)
    @article_number = article_number
    @href = "http://cgi.ebay.de/ws/eBayISAPI.dll?ViewItem&item=" + article_number
    @price = 
  end  
end

article_numbers.each do |article_number|
  p = agent.get("http://cgi.ebay.de/ws/eBayISAPI.dll?ViewItem&item=" + article_number)
  
  if p.body.scan("Erfolgreiches Gebot")
    d = p.search("//div[@class='item_description']").text # article description
    #price = d.search("span[@id='v4-31']").text.split.last.gsub(",", ".").to_f
    #title = d.search("h1[@class='vi-is1-titleH1']").text
    #subtitle = d.search("h1[@class='vi-is1-titleH2']").text
    gold = d.text.scan(/(\ Gold|g)?(\d+\.\d+)(\ Gold|g)?/) # 100g, Gold 100, 100 Gold

    # Kaltwetterflug = 375
    # Fliegen 250 = Epic Fliegen
    # Fliegen 225 = Nur Fliegen
    # Reiten 150 = Epic Reiten
    # Reiten 75 = Nur Reiten

    mount = d.scan(/(Kaltwetter\w+|[Rr]eiten|[Ff]liegen).?(\d{2,3})?/)
    dual = d.scan(/[Dd]ual[\w]?/)
    level = d.scan(/Stufe.{0,5}\d{2}/).first.split.last # stufe, lvl, level
    level = d.scan(/Stufe.?\d+|(\d+)(er|ig|iger)/)
    level = d.scan(/(Stufe.?(\d+))|(\d+)(er|ig|iger)/).flatten.compact.uniq # hier nur zahlen!
  else
    puts "failed"
  end
  break
end
  
  
  #auction = { :id => article_number, :name => String.new, :success => false, :price => 0}

  #auction[:id] = article_number
  #auction[:success] = true if p.body.scan("Erfolgreiches Gebot")
  #
  #print auction[:id] + " "
  #print "erfolgreich!" if auction[:success]
  #print "\n"
  #
  #
  #pp p.body.scan(/\w+dauer\w+/)