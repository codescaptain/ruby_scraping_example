require_relative 'dependencies'

class Main
  def scrape_page(opt1 = nil, opt2 = nil)
    Nokogiri::HTML(URI.open('https://siberci.com/'), opt1, opt2)
  end

  def get_html_data
    @siberci_data = []
    doc = scrape_page(nil, Encoding::UTF_8.to_s)
    parent_post = doc.css(".post-items li")
    parent_post.each do |the_post|
      post_title = the_post.css("article div.content header.entry-header h1.entry-title a").text
      @siberci_data << post_title
    end

    @siberci_data
  end

  def make_a_excel

    workbook = FastExcel.open("example.xlsx", constant_memory: true)

    workbook.default_format.set(
      font_size: 0, # user's default
    #font_family: "Arial"
      )
    
    worksheet = workbook.add_worksheet("Siberci.com Titles")

    bold = workbook.bold_format
    worksheet.set_column(0, 0, FastExcel::DEF_COL_WIDTH, bold)

    worksheet.append_row(["Title"], bold)

    get_html_data.each_with_index do |title, i|
      worksheet.append_row(["#{title}"])
    end

    workbook.close
    puts "Saved to file example.xlsx"
  end
end

puts Main.new.make_a_excel