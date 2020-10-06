require 'ruby_powerpoint'

files = Dir.glob("#{ENV["HOME"]}/**/*.pptx?")
files.each do |file|
  puts file
end
