#include "pch.h"
using namespace Spire::Doc;
using namespace Spire::Common;

int main() {
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"CreateTable.docx";

	//Open a blank Word document as template
	Document* document = new Document();

	Section* section = document->AddSection();
	addTable(section);

	//Save docx file
	document->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	document->Close();
	delete document;
}

void addTable(Section* section)
{
	vector<wstring> header = { L"Name", L"Capital", L"Continent", L"Area", L"Population" };
	vector<vector<wstring>> data =
	{
		{L"Argentina", L"Buenos Aires", L"South America", L"2777815", L"32300003"},
		{L"Bolivia", L"La Paz", L"South America", L"1098575", L"7300000"},
		{L"Brazil", L"Brasilia", L"South America", L"8511196", L"150400000"},
		{L"Canada", L"Ottawa", L"North America", L"9976147", L"26500000"},
		{L"Chile", L"Santiago", L"South America", L"756943", L"13200000"},
		{L"Colombia", L"Bagota", L"South America", L"1138907", L"33000000"},
		{L"Cuba", L"Havana", L"North America", L"114524", L"10600000"},
		{L"Ecuador", L"Quito", L"South America", L"455502", L"10600000"},
		{L"El Salvador", L"San Salvador", L"North America", L"20865", L"5300000"},
		{L"Guyana", L"Georgetown", L"South America", L"214969", L"800000"},
		{L"Jamaica", L"Kingston", L"North America", L"11424", L"2500000"},
		{L"Mexico", L"Mexico City", L"North America", L"1967180", L"88600000"},
		{L"Nicaragua", L"Managua", L"North America", L"139000", L"3900000"},
		{L"Paraguay", L"Asuncion", L"South America", L"406576", L"4660000"},
		{L"Peru", L"Lima", L"South America", L"1285215", L"21600000"},
		{L"United States of America", L"Washington", L"North America", L"9363130", L"249200000"},
		{L"Uruguay", L"Montevideo", L"South America", L"176140", L"3002000"},
		{L"Venezuela", L"Caracas", L"South America", L"912047", L"19700000"}
	};
	Table* table = section->AddTable(true);
	table->ResetCells(data.size() + 1, header.size());

	// ***************** First Row *************************
	TableRow* row = table->GetRows()->GetItem(0);
	row->SetIsHeader(true);
	row->SetHeight(20); //unit: point, 1point = 0.3528 mm
	row->SetHeightType(TableRowHeightType::Exactly);
	row->GetRowFormat()->SetBackColor(Color::GetGray());
	for (int i = 0; i < header.size(); i++)
	{
		row->GetCells()->GetItem(i)->GetCellFormat()->SetVerticalAlignment(VerticalAlignment::Middle);
		Paragraph* p = row->GetCells()->GetItem(i)->AddParagraph();
		p->GetFormat()->SetHorizontalAlignment(HorizontalAlignment::Center);
		TextRange* txtRange = p->AppendText(header[i].c_str());
		txtRange->GetCharacterFormat()->SetBold(true);
	}

	for (int r = 0; r < data.size(); r++)
	{
		TableRow* dataRow = table->GetRows()->GetItem(r + 1);
		dataRow->SetHeight(20);
		dataRow->SetHeightType(TableRowHeightType::Exactly);
		dataRow->GetRowFormat()->SetBackColor(Color::Empty());
		for (int c = 0; c < data[r].size(); c++)
		{
			dataRow->GetCells()->GetItem(c)->GetCellFormat()->SetVerticalAlignment(VerticalAlignment::Middle);
			dataRow->GetCells()->GetItem(c)->AddParagraph()->AppendText(data[r][c].c_str());
		}
	}

	for (int j = 1; j < table->GetRows()->GetCount(); j++)
	{
		if (j % 2 == 0)
		{
			TableRow* row2 = table->GetRows()->GetItem(j);
			for (int f = 0; f < row2->GetCells()->GetCount(); f++)
			{
				row2->GetCells()->GetItem(f)->GetCellFormat()->SetBackColor(Color::GetLightBlue());
			}
		}
	}
}