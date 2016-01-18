require 'spec_helper'

describe SpreadsheetDocument do
  context 'w/ filetype option' do
    subject { described_class.new(VALID_SPREADSHEET) }

    its(:filetype) { should eq('xls') }
  end

  context 'with filetype option' do
    subject { described_class.new(VALID_SPREADSHEET, filetype: 'xls') }

    its(:filetype) { should eq('xls') }
  end

  describe '#xlsx?' do
    let(:worksheet) { described_class.new 'Empty path' }
    before do
      worksheet.instance_variable_set(:@filetype, file_type)
    end

    context 'when xlsx document' do
      let(:file_type) { 'xlsx' }
      subject { worksheet.xlsx? }

      it { is_expected.to be_truthy }
    end

    context 'when xls document' do
      let(:file_type) { 'xls' }
      subject { worksheet.xlsx? }

      it { is_expected.to be_falsey }
    end
  end
end

