module Exchanger
  # The DeleteAttachment element is the root element in a request to delete an attachment from the Exchange store.
  #
  # https://msdn.microsoft.com/en-us/library/office/aa580612(v=exchg.150).aspx
  class DeleteAttachment < Operation
    class Request < Operation::Request
      attr_accessor :attachment_ids

      # Reset request options to defaults.
      def reset
        @attachment_ids = []
      end

      def to_xml
        super do |xml|
          xml.DeleteAttachment("xmlns" => NS["m"]) do
            xml.AttachmentIds do
              attachment_ids.each do |attachment_id|
                xml["t"].AttachmentId("Id" => attachment_id)
              end
            end
          end
        end
      end
    end

    class Response < Operation::Response
    end
  end
end
