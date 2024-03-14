package com.github.yhao3.excelpro.domain.syntax;

public abstract class CollapsibleSyntax extends Syntax {

    public static String getPrefix() {
        throw new UnsupportedOperationException("This method is not supported for CollapsibleSyntax");
    }

    public static String getSuffix() {
        throw new UnsupportedOperationException("This method is not supported for CollapsibleSyntax");
    }
}
