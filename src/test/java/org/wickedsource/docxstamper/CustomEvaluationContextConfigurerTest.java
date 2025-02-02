package org.wickedsource.docxstamper;

import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.wml.P;
import org.junit.Assert;
import org.junit.Test;
import org.springframework.expression.EvaluationContext;
import org.springframework.expression.PropertyAccessor;
import org.springframework.expression.TypedValue;
import org.wickedsource.docxstamper.util.ParagraphWrapper;

import java.io.IOException;
import java.io.InputStream;

public class CustomEvaluationContextConfigurerTest extends AbstractDocx4jTest {

  @Test
  public void customEvaluationContextConfigurerIsHonored() throws Docx4JException, IOException {
    DocxStamperConfiguration config = new DocxStamperConfiguration();
    config.setEvaluationContextConfigurer(context -> context.addPropertyAccessor(new SimpleGetter("foo", "bar")));

    InputStream template = getClass().getResourceAsStream("CustomEvaluationContextConfigurerTest.docx");
    WordprocessingMLPackage document = stampAndLoad(template, new EmptyContext(), config);

    P p2 = (P) document.getMainDocumentPart().getContent().get(2);
    Assert.assertEquals("The variable foo has the value bar.", new ParagraphWrapper(p2).getText());
  }

  static class EmptyContext {

  }

  static class SimpleGetter implements PropertyAccessor {

    private final String fieldName;

    private final Object value;

    public SimpleGetter(String fieldName, Object value) {
      this.fieldName = fieldName;
      this.value = value;
    }

    @Override
    public Class<?>[] getSpecificTargetClasses() {
      return null;
    }

    @Override
    public boolean canRead(EvaluationContext context, Object target, String name) {
      return true;
    }

    @Override
    public TypedValue read(EvaluationContext context, Object target, String name) {
      if (name.equals(this.fieldName)) {
        return new TypedValue(value);
      } else {
        return null;
      }
    }

    @Override
    public boolean canWrite(EvaluationContext context, Object target, String name) {
      return false;
    }

    @Override
    public void write(EvaluationContext context, Object target, String name, Object newValue) {

    }
  }
}
