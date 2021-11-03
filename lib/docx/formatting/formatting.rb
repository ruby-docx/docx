
# 給节点添加属性
module Docx
  module Formatting
    def add_property(tag)
      properties_node.remove if properties_node.at_xpath(".//w:#{tag}") # Remove and replace property
      properties_node.add_child("<w:#{tag}/>").first
    end

    def properties_node
      properties = node.at_xpath(".//w:#{properties_tag}")
      # Should a paragraph formatting node not exist create one
      properties ||= node.prepend_child("<w:#{properties_tag}></w:#{properties_tag}>").first
    end
  end
end
