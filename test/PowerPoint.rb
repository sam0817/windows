require 'win32ole' unless defined?(WIN32OLE)

class PowerPoint

  # initial PowerPoint Application
  def initialize
    begin
      @powerpoint=WIN32OLE.connect('Powerpoint.Application')
    rescue WIN32OLERuntimeError
      @powerpoint=WIN32OLE.new('Powerpoint.Application')
    end
  end

  def presentations
    @powerpoint.presentations
  end

  def presentation(index=1)
    @active_presentation=@powerpoint.presentations(index)
  end

  def slides
    @active_presentation.slides
  end

  def slide(index=1)
    @active_slide=@active_presentation.slides(index)
  end

  def slide_by_id(id)
    @active_slide=@active_presentation.slides.findbyslideid(id)
  end

  def shapes(index)
    @active_shape=@active_slide.shapes(index)
  end

  def textframe=(text)
    @active_shape.textframe.textrange.text=text unless @active_shape.hastextframe == 0
  end

  # 2-dimension array for tables [[row1],[row2]]
  def table=(array)
    unless @active_shape.hastable == 0
      array.each_with_index do |row_array, row_no|
        if @active_shape.table.rows.count > row_no
          row_array.each_with_index do |value, col_no|
            if @active_shape.table.columns.count > col_no
              @active_shape.table.cell(1 + row_no, 1 + col_no).shape.textframe.textrange.text = value unless value.nil?
            end
          end
        end
      end
    end
  end

  #TODO: need
  def chartdata=(array)
    #ActivePresentation.Slides(idx).Shapes(2).Chart.ChartData.Workbook.worksheets(1).listobjects(1).resize _
    #ppt.Slides.FindBySlideID(258).Shapes(2).Chart.ChartData.Activate
    #'ActiveSheet.ListObjects("Table1").Resize Range("$A$1:$D$4")
    #'ppt.Slides.FindBySlideID(258).Shapes(2).Chart.ChartData.workbook.worksheets(1).listobjects
    #Debug.Print ActivePresentation.Slides(idx).Shapes(2).Chart.ChartData.Workbook.worksheets(1).listobjects(1).Range.Address
    #ActivePresentation.Slides(idx).Shapes(2).Chart.ChartData.Workbook.worksheets(1).listobjects(1).resize _
    #ActivePresentation.Slides(idx).Shapes(2).Chart.ChartData.Workbook.worksheets(1).Range("A1:C4")
    #ActivePresentation.Slides(idx).Shapes(2).Chart.ChartData.Workbook.Close
  end
end

