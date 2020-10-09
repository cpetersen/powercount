require 'csv'
require 'ruby_powerpoint'

# files = Dir.glob("#{ENV["HOME"]}/**/*.ppt*")
files = Dir.glob("/Volumes/Dropbox/Scientist/**/*.pptx"); files.count
files.each do |file|
  # next unless file.match(/\.pptx?$/i)
  # next unless file.match(/\.pptx$/i)
  next if file.include?('~$')
  next if file.include?('[Autosaved]')
  begin
    deck = RubyPowerpoint::Presentation.new file
    puts CSV.generate_line([file, deck.slides.count])
  rescue => e
    puts CSV.generate_line([file, "Err", e.message])
  end
end;nil
