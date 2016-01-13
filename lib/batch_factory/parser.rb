module BatchFactory
  class Parser
    attr_accessor :worksheet, :user_heading_keys,
      :heading_keys, :column_bounds, :row_hashes

    def initialize
      @heading_keys = []
      @column_bounds = []
      @row_hashes = []
    end

    def open(file_location, options = {})
      @worksheet = SpreadsheetDocument.new(file_location, options)
      @user_heading_keys = options[:keys]
    end

    def parse!
      @hashed_worksheet = nil

      parse_column_bounds
      parse_heading_keys
      parse_data_rows
    end

    def parse_heading_keys
      @heading_keys = user_heading_keys || worksheet.row(worksheet.first_row).map do |key|
        key.blank? ? nil : key.strip
      end
    end

    def parse_column_bounds
      @column_bounds = (@worksheet.first_column-1)..(@worksheet.last_column-1)
    end

    def parse_data_rows
      @row_hashes = []

      rows_range.each do |row_idx|
        @row = @worksheet.row(row_idx)
        hash = HashWithIndifferentAccess.new
        @column_bounds.each do |cell_idx|
          value = cell_value(row_idx, cell_idx)
          if key = @heading_keys[cell_idx]
            hash[key] = value if value.present?
          end
        end
        @row_hashes << hash unless hash.empty?
      end
    end

    def hashed_worksheet
      @hashed_worksheet ||= HashedWorksheet.new(
        @heading_keys,
        @row_hashes
      )
    end

    private

    def numeric_cell?(row_idx, cell_idx)
      @worksheet.excelx_type(row_idx, cell_idx) == [:numeric_or_formula, 'General']
    end

    def first_row_offset
      @user_heading_keys ? 0 : 1
    end

    def rows_range
      (@worksheet.first_row + first_row_offset)..@worksheet.last_row
    end

    def cell_value(row_idx, cell_idx)
      roo_cell_idx = cell_idx + 1 # cells indexes in worksheet row start with 1
      if @worksheet.xlsx? and numeric_cell?(row_idx, roo_cell_idx)
        @worksheet.excelx_value(row_idx, roo_cell_idx)
      else
        @row[cell_idx]
      end
    end

  end
end
