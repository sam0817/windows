require_relative '../test/PowerPoint'

ppt=PowerPoint.new

puts ppt.presentation.name
puts ppt.slides.count
puts ppt.slide_by_id(257)
puts ppt.shapes(2).ole_type
ppt.textframe = 'Sam\'s 3rd test'
ppt.table =[[1, 2, 3, nil, 5], [2, 3, 4, 5, 6, 7]]

