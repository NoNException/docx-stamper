package org.wickedsource.docxstamper;

import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.wml.P;
import org.junit.Assert;
import org.junit.Test;
import org.wickedsource.docxstamper.util.ParagraphWrapper;

import java.io.IOException;
import java.io.InputStream;

public class ExpressionWithSurroundingSpacesTest extends AbstractDocx4jTest {

    @Test
    public void test() throws Docx4JException, IOException {
        Context context = new Context();
        InputStream template = getClass().getResourceAsStream("ExpressionWithSurroundingSpacesTest.docx");
        WordprocessingMLPackage document = stampAndLoad(template, context);

        Assert.assertEquals("Before Expression After.", new ParagraphWrapper((P) document.getMainDocumentPart().getContent().get(2)).getText());
        Assert.assertEquals("Before Expression After.", new ParagraphWrapper((P) document.getMainDocumentPart().getContent().get(3)).getText());
        Assert.assertEquals("Before Expression After.", new ParagraphWrapper((P) document.getMainDocumentPart().getContent().get(4)).getText());
        Assert.assertEquals("Before Expression After.", new ParagraphWrapper((P) document.getMainDocumentPart().getContent().get(5)).getText());
        Assert.assertEquals("Before Expression After.", new ParagraphWrapper((P) document.getMainDocumentPart().getContent().get(6)).getText());
        Assert.assertEquals("Before Expression After.", new ParagraphWrapper((P) document.getMainDocumentPart().getContent().get(7)).getText());
        Assert.assertEquals("Before Expression After.", new ParagraphWrapper((P) document.getMainDocumentPart().getContent().get(8)).getText());
    }

    static class Context {
        private final String expressionWithLeadingAndTrailingSpace = " Expression ";
        private final String expressionWithLeadingSpace = " Expression";
        private final String expressionWithTrailingSpace = "Expression ";
        private final String expressionWithoutSpaces = "Expression";

        public String getExpressionWithLeadingAndTrailingSpace() {
            return " Expression ";
        }

        public String getExpressionWithLeadingSpace() {
            return " Expression";
        }

        public String getExpressionWithTrailingSpace() {
            return "Expression ";
        }

        public String getExpressionWithoutSpaces() {
            return "Expression";
        }
    }


}
