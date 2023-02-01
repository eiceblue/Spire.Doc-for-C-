#include "pch.h"
using namespace Spire::Doc;
using namespace Spire::Common;

int main() {
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"PageSetup.doc";

	//Create Word document.
	Document* document = new Document();
	Section* section = document->AddSection();

	//The unit of all measures below is point, 1point = 0.3528 mm.
	section->GetPageSetup()->SetPageSize(PageSize::A4());
	section->GetPageSetup()->GetMargins()->SetTop(72.0f);
	section->GetPageSetup()->GetMargins()->SetBottom(72.0f);
	section->GetPageSetup()->GetMargins()->SetLeft(89.85f);
	section->GetPageSetup()->GetMargins()->SetRight(89.85f);

	//Insert header and footer.
	InsertHeaderAndFooter(section);

	addTable(section);

	//Save doc file.
	document->SaveToFile(outputFile.c_str(), FileFormat::Doc);
	document->Close();
	delete document;
}

void addTable(Section* section)
{
	vector<wstring> header = { L"Name", L"Capital", L"Continent", L"Area", L"Population" };
	vector<vector<wstring>> data = {
		{L"Argentina", L"Buenos Aires", L"South America", L"2777815", L"32300003"},
		{L"Bolivia", L"La Paz", L"South", L"1098575", L"7300000"},
		{L"Brazil", L"Brasilia", L"South", L"8511196", L"150400000"},
		{L"Canada", L"Ottawa", L"North", L"9976147", L"26500000"},
		{L"Chile", L"Santiago", L"South", L"756943", L"13200000"},
		{L"Colombia", L"Bagota", L"South", L"1138907", L"33000000"},
		{L"Cuba", L"Havana", L"North", L"114524", L"10600000"},
		{L"Ecuador", L"Quito", L"South", L"455502", L"10600000"},
		{L"El Salvador", L"San Salvador", L"North", L"20865", L"5300000"},
		{L"Guyana", L"Georgetown", L"South", L"214969", L"800000"},
		{L"Jamaica", L"Kingston", L"North", L"11424", L"2500000"},
		{L"Mexico", L"Mexico City", L"North", L"1967180", L"88600000"},
		{L"Nicaragua", L"Managua", L"North", L"139000", L"3900000"},
		{L"Paraguay", L"Asuncion", L"South", L"406576", L"4660000"},
		{L"Peru", L"Lima", L"South", L"1285215", L"21600000"},
		{L"United States", L"Washington", L"North", L"9363130", L"249200000"},
		{L"Uruguay", L"Montevideo", L"South", L"176140", L"3002000"},
		{L"Venezuela", L"Caracas", L"South", L"912047", L"19700000"}
	};
	Table* table = section->AddTable(true);

	table->ResetCells(data.size() + 1, header.size());

	// ***************** First Row *************************
	TableRow* row = table->GetRows()->GetItem(0);
	row->SetIsHeader(true);
	row->SetHeight(20);
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
}

void InsertHeaderAndFooter(Section* section)
{
	wstring input_path = DATAPATH;
	HeaderFooter* header = section->GetHeadersFooters()->GetHeader();
	HeaderFooter* footer = section->GetHeadersFooters()->GetFooter();

	//Insert picture and text to header.
	Paragraph* headerParagraph = header->AddParagraph();
	wstring imagePath1 = input_path + L"Header.png";
	DocPicture* headerPicture = headerParagraph->AppendPicture(imagePath1.c_str());

	//Header text.
	TextRange* text = headerParagraph->AppendText(L"Demo of Spire.Doc");
	text->GetCharacterFormat()->SetFontName(L"Arial");
	text->GetCharacterFormat()->SetFontSize(10);
	text->GetCharacterFormat()->SetItalic(true);
	headerParagraph->GetFormat()->SetHorizontalAlignment(HorizontalAlignment::Right);

	//Border.
	headerParagraph->GetFormat()->GetBorders()->GetBottom()->SetBorderType(BorderStyle::Single);
	headerParagraph->GetFormat()->GetBorders()->GetBottom()->SetSpace(0.05F);


	//Header picture layout - text wrapping.
	headerPicture->SetTextWrappingStyle(TextWrappingStyle::Behind);

	//Header picture layout - position.
	headerPicture->SetHorizontalOrigin(HorizontalOrigin::Page);
	headerPicture->SetHorizontalAlignment(ShapeHorizontalAlignment::Left);
	headerPicture->SetVerticalOrigin(VerticalOrigin::Page);
	headerPicture->SetVerticalAlignment(ShapeVerticalAlignment::Top);

	//Insert picture to footer.
	Paragraph* footerParagraph = footer->AddParagraph();
	wstring imagePath2 = input_path + L"Footer.png";
	DocPicture* footerPicture = footerParagraph->AppendPicture(imagePath2.c_str());

	//Footer picture layout.
	footerPicture->SetTextWrappingStyle(TextWrappingStyle::Behind);
	footerPicture->SetHorizontalOrigin(HorizontalOrigin::Page);
	footerPicture->SetHorizontalAlignment(ShapeHorizontalAlignment::Left);
	footerPicture->SetVerticalOrigin(VerticalOrigin::Page);
	footerPicture->SetVerticalAlignment(ShapeVerticalAlignment::Bottom);

	//Insert page number.
	footerParagraph->AppendField(L"page number", FieldType::FieldPage);
	footerParagraph->AppendText(L" of ");
	footerParagraph->AppendField(L"number of pages", FieldType::FieldNumPages);
	footerParagraph->GetFormat()->SetHorizontalAlignment(HorizontalAlignment::Right);

	//Border.
	footerParagraph->GetFormat()->GetBorders()->GetTop()->SetBorderType(BorderStyle::Single);
	footerParagraph->GetFormat()->GetBorders()->GetTop()->SetSpace(0.05F);
}
