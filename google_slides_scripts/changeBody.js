// This script changes body text settings (everything except TITLES) to Helvetica Neue 20pt
function changeBodyText() {
  const presentation = SlidesApp.getActivePresentation();
  const slides = presentation.getSlides();
  let changedCount = 0;

  slides.forEach(slide => {
    const titlePlaceholder = slide.getPlaceholder(SlidesApp.PlaceholderType.TITLE);
    const pageElements = slide.getPageElements();
    
    pageElements.forEach(element => {
      if (element.getPageElementType() !== SlidesApp.PageElementType.SHAPE) return;
      // Check whether an element has type = TITLE; skip it
      if (titlePlaceholder && element.getObjectId() === titlePlaceholder.getObjectId()) return;
      // Otherwise, choose parameters to change in the elem
      const shape = element.asShape();
      const textRange = shape.getText();
      
      if (!textRange || textRange.asString().trim() === "") return;
      
      textRange.getTextStyle()
                .setFontFamily("Helvetica Neue") // choose font
                .setFontSize(20) // choose font size in pt
                .setForegroundColor("#000000"); // choose font color in hex
      changedCount++;
    });
  });
  
  console.log(`Changed body text elements: ${changedCount}`);
}
