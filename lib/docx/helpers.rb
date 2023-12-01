module Docx
  module SimpleInspect
    # Returns a string representation of the document that is far more readable and understandable
    # than the default inspect method. But you can still get the default inspect method by passing
    # true as the first argument.
    def inspect(full = false)
      return(super) if full

      variable_values =
        instance_variables.map do |var|
          value = v = instance_variable_get(var).inspect

          [
            var,
            value.length > 100 ? "#{value[0..100]}..." : value
          ].join('=')
        end

      "#<#{self.class}:0x#{(object_id << 1).to_s(16)} #{variable_values.join(' ')}>"
    end
  end
end