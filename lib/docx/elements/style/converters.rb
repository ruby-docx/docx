module Docx
  module Elements
    class Style
      module Converters
        class DefaultValueConverter
          def self.encode(value)
            value
          end

          def self.decode(value)
            value
          end
        end

        class FontSizeConverter
          def self.encode(value)
            value.to_i * 2
          end

          def self.decode(value)
            value.to_i / 2
          end
        end

        class BooleanConverter
          def self.encode(value)
            value ? "1" : "0"
          end

          def self.decode(value)
            value == "1"
          end
        end
      end
    end
  end
end
