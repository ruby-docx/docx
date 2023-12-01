module Docx
  module Elements
    class Style
      module Validators
        class DefaultValidator
          def self.validate(value)
            true
          end
        end

        class ColorValidator
          COLOR_REGEX = /^([A-Fa-f0-9]{6}|[A-Fa-f0-9]{3})$/

          def self.validate(value)
            value =~ COLOR_REGEX
          end
        end
      end
    end
  end
end
