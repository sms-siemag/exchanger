module Exchanger
  # The FindItem operation identifies items that are located in a specified folder.
  #
  # http://msdn.microsoft.com/en-us/library/aa566107.aspx
  class FindItem < Operation
    class Request < Operation::Request
      attr_accessor :folder_id, :traversal, :base_shape, :email_address, :calendar_view, :max_entries_returned, :offset, :base_point

      # Reset request options to defaults.
      def reset
        @folder_id = :contacts
        @traversal = :shallow
        @base_shape = :all_properties
        @email_address = nil
        @calendar_view = nil
        @max_entries_returned = nil
        @offset = 0
        @base_point = :Beginning
      end

      def to_xml
        super do |xml|
          xml.FindItem("xmlns" => NS["m"], "xmlns:t" => NS["t"], "Traversal" => traversal.to_s.camelize) do
            xml.ItemShape do
              xml.send "t:BaseShape", base_shape.to_s.camelize
            end
            if calendar_view
              xml.CalendarView(calendar_view.to_xml.attributes)
            end
            if max_entries_returned.present? or offset.present? or base_point.present?
              xml.IndexedPageItemView(indexed_page_item_view)
            end
            xml.ParentFolderIds do
              if folder_id.is_a?(Symbol)
                xml.send("t:DistinguishedFolderId", "Id" => folder_id) do
                  if email_address
                    xml.send("t:Mailbox") do
                      xml.send("t:EmailAddress", email_address)
                    end
                  end
                end
              else
                xml.send("t:FolderId", "Id" => folder_id)
              end
            end
          end
        end
      end

    private
      def indexed_page_item_view
        hItem = {'Offset' => offset, 'BasePoint' => base_point}
        hItem['MaxEntriesReturned'] = max_entries_returned if max_entries_returned

        hItem
      end
    end

    class Response < Operation::Response
      def items
        to_xml.xpath(".//t:Items", NS).children.map do |node|
          item_klass = Exchanger.const_get(node.name)
          item_klass.new_from_xml(node)
        end
      end
    end
  end
end
