require "optparse"
require "rubygems"
require "tmpdir"
require "pathname"
require "rexml/document"

class MD2XML
  def add_header(text, level, body)
    @in_list = false
    node = REXML::Element.new("w:p", body)
    pPr = REXML::Element.new("w:pPr", node)
    REXML::Element.new("w:pStyle", pPr).add_attribute("w:val", level)
    r = REXML::Element.new("w:r", node)
    REXML::Element.new("w:t", r).text = text
  end

  def add_listitem(text, level, style, body)
    unless @in_list
      @in_list = true
      @id += 1
      @list_id[level] = @id
    end
    if level > @current_level
      @id += 1
      @list_id[level] = @id
    end
    @current_level = level
    myid = @list_id[level]
    node = REXML::Element.new("w:p", body)
    pPr = REXML::Element.new("w:pPr", node)
    pStyle = REXML::Element.new("w:pStyle", pPr)
    numPr = REXML::Element.new("w:numPr", pPr)
    REXML::Element.new("w:ilvl", numPr).add_attribute("w:val", (level-1).to_s)
    REXML::Element.new("w:numId", numPr).add_attribute("w:val", myid)
    pstyle_val = REXML::XPath.first(@stylehash[style], "w:pPr/w:pStyle").attribute("w:val")
    pStyle.add_attribute("w:val", pstyle_val)
    REXML::Element.new("w:ind", pPr).add_attribute("w:leftChars", "0")
    r = REXML::Element.new("w:r", node)
    REXML::Element.new("w:t", r).text = text
    numid = REXML::XPath.first(@stylehash[style], "w:pPr/w:numPr/w:numId").attribute("w:val").to_s
    @numIdhash[myid] = @numhash[numid]
  end

  def add_paragraph(text, body)
    @in_list = false
    p = REXML::Element.new("w:p", body)
    r = REXML::Element.new("w:r", p)
    REXML::Element.new("w:t", r).text = text
  end

  def parse(line, body)
    if line=~/^(#+) (.*)/
      add_header(Regexp.last_match(2), Regexp.last_match(1).size, body)
    elsif line=~/^(\s*)[0-9]+\. (.*)/
      level = Regexp.last_match(1).length/4+1
      add_listitem(Regexp.last_match(2), level, "enum"+level.to_s, body)
    elsif line=~/^(\s*)\* (.*)/
      level = Regexp.last_match(1).length/4+1
      add_listitem(Regexp.last_match(2), level, "bullet"+level.to_s, body)
    else
      add_paragraph(line, body)
    end
  end

  def make_numhash(dir)
    @numhash = {}
    file = dir + "/word/numbering.xml"
    doc = REXML::Document.new(File.read(file))
    REXML::XPath.each(doc.root, "w:num") do |e|
      numid = e.attribute("w:numId").to_s.to_i
      abstractnumid = REXML::XPath.first(e, "w:abstractNumId").attribute("w:val")
      @numhash[numid.to_s] = abstractnumid.to_s.to_i
      @id = numid if @id < numid
    end
  end

  def make_numbering(dir)
    file = dir + "/word/numbering.xml"
    doc = REXML::Document.new(File.read(file))
    abstractNumIds = REXML::XPath.each(doc.root, "w:abstractNum").collect { |e| e }
    nums = REXML::XPath.each(doc.root, "w:num").collect { |e| e }
    REXML::XPath.first(doc.root).each { |e| doc.root.delete e }

    @numIdhash.each do |k, v|
      e_abs = Marshal.load(Marshal.dump(abstractNumIds[v]))
      e_abs.add_attribute("w:abstractNumId", abstractNumIds.size)
      n = REXML::Element.new("w:num")
      n.add_attribute("w:numId", k)
      REXML::Element.new("w:abstractNumId", n).add_attribute("w:val", abstractNumIds.size)
      abstractNumIds.push e_abs
      nums.push n
    end

    abstractNumIds.size.times do |i|
      e = abstractNumIds[i]
      REXML::XPath.first(e, "w:nsid").add_attribute("w:val", format("%08d", i))
      doc.root.add e
    end

    nums.each { |e| doc.root.add e }
    File.write file, doc.to_s
  end

  def make_document(dir, mdfile)
    file = dir + "/word/document.xml"
    doc = REXML::Document.new(File.read(file))
    body = REXML::XPath.first(doc.root, "w:body")
    REXML::XPath.each(body, "w:p") do |e|
      @stylehash[Regexp.last_match(1)] = e if e.to_s =~/(enum[1-9])/ || e.to_s =~/(bullet[1-9])/
    end
    REXML::XPath.each(body, "w:p").collect { |e| body.delete_element e }

    open(mdfile) do |f|
      while line = f.gets
        parse(line, body)
      end
    end

    File.write file, doc.to_s
  end

  def convert(dir, mdfile)
    @id = 0
    make_numhash(dir)
    @in_list = false
    @list_id = Array.new(10)
    @stylehash = {}
    @current_level = 0
    @numIdhash = {}
    make_document(dir, mdfile)
    make_numbering(dir)
  end
end

def parse_option
  args = {}
  OptionParser.new do |op|
    op.on("-i [input]", "--input [input file]") { |v| args[:input] = v }
    op.on("-t [template]", "--template [template file]") { |v| args[:template] = v }
    op.on("-o [output]", "--output [output file]") { |v| args[:output] = v }
    op.parse!(ARGV)
  end
  args
end

args = parse_option
input_file = args.fetch(:input, "input.md")
template_file = args.fetch(:template, "template.docx")

output_file = args.fetch(:output, "output.docx")

Dir.mktmpdir(nil, "./") do |dir|
  puts "Using #{template_file}"
  `cd #{dir};unzip ../#{template_file}`
  files = Dir.glob(dir+"/*").map { |f| File.basename(f) }
  puts "Reading #{input_file}"
  MD2XML.new.convert(dir, input_file)
  puts "Generating #{output_file}"
  `cd #{dir};zip -r ../#{output_file} #{files.join(" ")}`
  puts "Done."
end
