// This script changes title settings, e.g. all titles to Helvetica Bold 30pt
// To run, open Extentions > Apps Scrips > copy, paste and run this guy 
function changeOnlySlideTitles() {
  const presentation = SlidesApp.getActivePresentation();
  const slides = presentation.getSlides();
  let changedCount = 0;

  slides.forEach(slide => {
    const titlePlaceholder = slide.getPlaceholder(SlidesApp.PlaceholderType.TITLE);
    
    if (titlePlaceholder) {
      const textRange = titlePlaceholder.asShape().getText();
      if (textRange && textRange.asString().trim() !== "") {
        textRange.getTextStyle()
                  .setFontFamily("Helvetica Bold") // Change all titles to lovely Helvetica Bold; you can use your own fonts
                  .setFontSize(30) // Change all fonts to 30pt; change this if needed
                  .setForegroundColor("#000000"); // Switch to brilliant black color; also, change is needed
        changedCount++;
      }
    }
  });
  
  console.log(`Changed titles: ${changedCount}`);
}
