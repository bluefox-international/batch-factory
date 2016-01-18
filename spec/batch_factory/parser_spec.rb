require 'spec_helper'

describe BatchFactory::Parser do
  let(:parser) { BatchFactory::Parser.new }

  context 'w/ instance methods' do
    it 'should open and parse a valid file' do
      parser.open VALID_SPREADSHEET
      expect(parser.worksheet).not_to be_nil
    end

    context 'when parsing default worksheet' do
      before do
        parser.open VALID_SPREADSHEET
        parser.parse!
      end

      it 'should parse column information' do
        expect(parser.column_bounds).to eq(0..5)
      end

      it 'should parse the heading keys' do
        expect(parser.heading_keys).to eq(['name', 'address', nil, 'country', 'bid', 'age'])
      end

      it 'should parse the data rows' do
        expect(parser.row_hashes[0]['age']).to eq 50
        expect(parser.row_hashes.size).to eq 1
      end

      it 'should return a hashed worksheet' do
        worksheet = parser.hashed_worksheet
        expect(worksheet.rows).to eq(parser.row_hashes)
        expect(worksheet.keys).to eq(parser.heading_keys)
      end
    end

    context 'when parsing second worksheet' do
      before do
        parser.open VALID_SPREADSHEET, sheet_number: 1
        parser.parse!
      end

      it 'should parse column information' do
        expect(parser.column_bounds).to eq(2..7)
      end

      it 'should parse the heading keys' do
        expect(parser.heading_keys).to eq(['name', 'address', nil, 'country', 'bid', 'age'])
      end

      it 'should parse the data rows' do
        expect(parser.row_hashes[0]['age']).to eq 50
        expect(parser.row_hashes.size).to eq 1
      end

      it 'should return a hashed worksheet' do
        worksheet = parser.hashed_worksheet
        expect(worksheet.rows).to eq(parser.row_hashes)
        expect(worksheet.keys).to eq(parser.heading_keys)
      end
    end
  end

  describe '#numeric_cell?' do
    let(:worksheet) { double(:worksheet, excelx_type: excelx_type) }
    before do
      parser.instance_variable_set(:@worksheet, worksheet)
    end
    subject { parser.send(:numeric_cell?, 1, 1) }

    context 'when numeric cell' do
      let(:excelx_type) { [:numeric_or_formula, 'General'] }

      it { is_expected.to be_truthy }
    end

    context 'when not numeric cell' do
      let(:excelx_type) { :string }

      it { is_expected.to be_falsey }
    end
  end

  describe '#first_row_offset' do
    before do
      parser.instance_variable_set(:@user_heading_keys, keys)
    end
    subject { parser.send(:first_row_offset) }

    context 'with user heading keys' do
      let(:keys) { ['First column', 'Second column'] }

      it 'should equal 0' do
        expect(subject).to eq 0
      end
    end

    context 'without user heading keys' do
      let(:keys) { nil }

      it 'should equal 1' do
        expect(subject).to eq 1
      end
    end
  end

  describe '#rows_range' do
    let(:worksheet) { double(:worksheet, first_row: 1, last_row: 5) }
    before do
      parser.instance_variable_set(:@worksheet, worksheet)
      parser.instance_variable_set(:@user_heading_keys, keys)
    end
    subject { parser.send(:rows_range) }

    context 'with first row offset' do
      let(:keys) { ['First column', 'Second column'] }

      it 'should not be offset' do
        expect(subject).to eq 1..5
      end
    end

    context 'without first row offset' do
      let(:keys) { nil }

      it 'should be offset by 1' do
        expect(subject).to eq 2..5
      end
    end
  end

  describe '#cell_value' do
    let(:xlsx) { true }
    let(:numeric_cell) { true }
    let(:row) { ['value 1', 'value 2'] }
    let(:worksheet) { double(:worksheet, xlsx?: xlsx) }
    before do
      allow(parser).to receive(:numeric_cell?).and_return numeric_cell
      parser.instance_variable_set(:@worksheet, worksheet)
    end

    context 'when numeric cell of xlsx document' do
      let(:excelx_value) { 'cell value' }
      before do
        allow(parser.instance_variable_get(:@worksheet)).to receive(:excelx_value).with(1, 2).and_return excelx_value
      end
      subject { parser.send(:cell_value, 1, 1) }

      it 'should call worksheet excelx_value method' do
        expect(parser.instance_variable_get(:@worksheet)).to receive(:excelx_value).with(1, 2)
        subject
      end

      it 'should return correct value' do
        expect(subject).to eq excelx_value
      end
    end

    context 'when not xlsx document' do
      let(:xlsx) { false }
      before do
        parser.instance_variable_set(:@row, row)
      end
      subject { parser.send(:cell_value, 1, 0) }

      it 'should not call worksheet excelx_value method' do
        expect(parser.instance_variable_get(:@worksheet)).not_to receive(:excelx_value)
        subject
      end

      it 'should equal first cell of row' do
        expect(subject).to eq 'value 1'
      end
    end

    context 'when not numeric cell' do
      let(:numeric_cell) { false }
      before do
        parser.instance_variable_set(:@row, row)
      end
      subject { parser.send(:cell_value, 1, 0) }

      it 'should not call worksheet excelx_value method' do
        expect(parser.instance_variable_get(:@worksheet)).not_to receive(:excelx_value)
        subject
      end

      it 'should equal first cell of row' do
        expect(subject).to eq 'value 1'
      end
    end
  end

end

