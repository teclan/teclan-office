package com.teclan.office;

import java.util.Map;

public interface WordHandler {

    public boolean handle(String templatePath, Map<String,Object> content, String outputFile);
}
