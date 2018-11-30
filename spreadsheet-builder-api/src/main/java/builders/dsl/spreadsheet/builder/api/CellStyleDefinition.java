package builders.dsl.spreadsheet.builder.api;

import builders.dsl.spreadsheet.api.*;

import java.util.function.Consumer;

public interface CellStyleDefinition extends ForegroundFillProvider, BorderPositionProvider, ColorProvider {

    CellStyleDefinition base(String stylename);

    CellStyleDefinition background(String hexColor);
    CellStyleDefinition background(Color color);

    CellStyleDefinition foreground(String hexColor);
    CellStyleDefinition foreground(Color color);

    CellStyleDefinition fill(ForegroundFill fill);

    CellStyleDefinition font(Consumer<FontDefinition> fontConfiguration);

    /**
     * Sets the indent of the cell in spaces.
     * @param indent the indent of the cell in spaces
     */
    CellStyleDefinition indent(int indent);

    /**
     * Enables word wrapping
     *
     * @param text keyword
     */
    CellStyleDefinition wrap(Keywords.Text text);

    /**
     * Sets the rotation from 0 to 180 (flipped).
     * @param rotation the rotation from 0 to 180 (flipped)
     */
    CellStyleDefinition rotation(int rotation);

    CellStyleDefinition format(String format);

    CellStyleDefinition align(Keywords.VerticalAlignment verticalAlignment, Keywords.HorizontalAlignment horizontalAlignment);

    /**
     * Configures all the borders of the cell.
     * @param borderConfiguration border configuration
     */
    CellStyleDefinition border(Consumer<BorderDefinition> borderConfiguration);

    /**
     * Configures one border of the cell.
     * @param location border to be configured
     * @param borderConfiguration border configuration
     */
    CellStyleDefinition border(Keywords.BorderSide location, Consumer<BorderDefinition> borderConfiguration);

    /**
     * Configures two borders of the cell.
     * @param first first border to be configured
     * @param second second border to be configured
     * @param borderConfiguration border configuration
     */
    CellStyleDefinition border(Keywords.BorderSide first, Keywords.BorderSide second, Consumer<BorderDefinition> borderConfiguration);

    /**
     * Configures three borders of the cell.
     * @param first first border to be configured
     * @param second second border to be configured
     * @param third third border to be configured
     * @param borderConfiguration border configuration
     */
    CellStyleDefinition border(Keywords.BorderSide first, Keywords.BorderSide second, Keywords.BorderSide third, Consumer<BorderDefinition> borderConfiguration);

}
