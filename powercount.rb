require 'csv'
require 'ruby_powerpoint'

files = Dir.glob("#{ENV["HOME"]}/**/*.ppt*")
files.each do |file|
  # next unless file.match(/\.pptx?$/i)
  next unless file.match(/\.pptx$/i)
  next if file.include?('~$')
  next if file.include?('[Autosaved]')
  deck = RubyPowerpoint::Presentation.new file
  puts CSV.generate_line([file, deck.slides.count])
end;nil
