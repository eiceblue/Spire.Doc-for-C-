#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"ConvertedTemplate.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"EmbedAllFontsInPDF.pdf";

	Document* document = new Document();
	document->LoadFromFile(inputFile.c_str());
	//embeds full fonts by default when IsEmbeddedAllFonts is set to true.
	ToPdfParameterList* ppl = new ToPdfParameterList();
	ppl->SetIsEmbeddedAllFonts(true);

	//Save doc file to pdf.
	document->SaveToFile(outputFile.c_str(), ppl);
	document->Close();
	delete document;
}
