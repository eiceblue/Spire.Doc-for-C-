#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"Textbox.docx";

	//Create a Word document and and a section.
	Document* document = new Document();
	Section* section = document->AddSection();

	InsertTextbox(section);

	//Save docx file.
	document->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	document->Close();
	delete document;
}

void InsertTextbox(Section* section)
{
	Paragraph* paragraph = section->GetParagraphs()->GetCount() > 0 ? section->GetParagraphs()->GetItem(0) : section->AddParagraph();
	paragraph = section->AddParagraph();
	paragraph = section->AddParagraph();

	//Insert and format the first textbox.
	TextBox* textBox1 = paragraph->AppendTextBox(240, 35);
	textBox1->GetFormat()->SetHorizontalAlignment(ShapeHorizontalAlignment::Left);
	textBox1->GetFormat()->SetLineColor(Spire::Common::Color::GetGray());
	textBox1->GetFormat()->SetLineStyle(TextBoxLineStyle::Simple);
	textBox1->GetFormat()->SetFillColor(Spire::Common::Color::GetDarkSeaGreen());
	Paragraph* para = textBox1->GetBody()->AddParagraph();
	TextRange* txtrg = para->AppendText(L"Textbox 1 in the document");
	txtrg->GetCharacterFormat()->SetFontName(L"Lucida Sans Unicode");
	txtrg->GetCharacterFormat()->SetFontSize(14);
	txtrg->GetCharacterFormat()->SetTextColor(Spire::Common::Color::GetWhite());
	para->GetFormat()->SetHorizontalAlignment(HorizontalAlignment::Center);

	//Insert and format the second textbox.
	paragraph = section->AddParagraph();
	paragraph = section->AddParagraph();
	paragraph = section->AddParagraph();
	paragraph = section->AddParagraph();
	TextBox* textBox2 = paragraph->AppendTextBox(240, 35);
	textBox2->GetFormat()->SetHorizontalAlignment(ShapeHorizontalAlignment::Left);
	textBox2->GetFormat()->SetLineColor(Spire::Common::Color::GetTomato());
	textBox2->GetFormat()->SetLineStyle(TextBoxLineStyle::ThinThick);
	textBox2->GetFormat()->SetFillColor(Spire::Common::Color::GetBlue());
	textBox2->GetFormat()->SetLineDashing(LineDashing::Dot);
	para = textBox2->GetBody()->AddParagraph();
	txtrg = para->AppendText(L"Textbox 2 in the document");
	txtrg->GetCharacterFormat()->SetFontName(L"Lucida Sans Unicode");
	txtrg->GetCharacterFormat()->SetFontSize(14);
	txtrg->GetCharacterFormat()->SetTextColor(Spire::Common::Color::GetPink());
	para->GetFormat()->SetHorizontalAlignment(HorizontalAlignment::Center);

	//Insert and format the third textbox.
	paragraph = section->AddParagraph();
	paragraph = section->AddParagraph();
	paragraph = section->AddParagraph();
	paragraph = section->AddParagraph();
	TextBox* textBox3 = paragraph->AppendTextBox(240, 35);
	textBox3->GetFormat()->SetHorizontalAlignment(ShapeHorizontalAlignment::Left);
	textBox3->GetFormat()->SetLineColor(Spire::Common::Color::GetViolet());
	textBox3->GetFormat()->SetLineStyle(TextBoxLineStyle::Triple);
	textBox3->GetFormat()->SetFillColor(Spire::Common::Color::GetPink());
	textBox3->GetFormat()->SetLineDashing(LineDashing::DashDotDot);
	para = textBox3->GetBody()->AddParagraph();
	txtrg = para->AppendText(L"Textbox 3 in the document");
	txtrg->GetCharacterFormat()->SetFontName(L"Lucida Sans Unicode");
	txtrg->GetCharacterFormat()->SetFontSize(14);
	txtrg->GetCharacterFormat()->SetTextColor(Spire::Common::Color::GetTomato());
	para->GetFormat()->SetHorizontalAlignment(HorizontalAlignment::Center);
}
