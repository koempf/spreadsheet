/*
 * SPDX-License-Identifier: Apache-2.0
 *
 * Copyright 2020 Vladimir Orany.
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     https://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */
package builders.dsl.spreadsheet.builder.poi

import builders.dsl.spreadsheet.builder.api.SpreadsheetBuilder
import builders.dsl.spreadsheet.builder.tck.AbstractBuilderSpec
import builders.dsl.spreadsheet.query.api.SpreadsheetCriteria
import builders.dsl.spreadsheet.query.poi.PoiSpreadsheetCriteria
import org.junit.Rule
import org.junit.rules.TemporaryFolder

import java.awt.Desktop

class PoiStreamingExcelBuilderSpec extends AbstractBuilderSpec {

    @Rule TemporaryFolder tmp = new TemporaryFolder()

    File tmpFile

    void setup() {
        tmpFile = tmp.newFile("sample${System.currentTimeMillis()}.xlsx")
    }

    @Override
    protected SpreadsheetCriteria createCriteria() {
        return PoiSpreadsheetCriteria.FACTORY.forFile(tmpFile)
    }

    @Override
    protected SpreadsheetBuilder createSpreadsheetBuilder() {
        return PoiSpreadsheetBuilder.stream(tmpFile)
    }

    @Override
    protected void openSpreadsheet() {
        open tmpFile
    }

    /**
     * Tries to open the file in Word. Only works locally on Mac at the moment. Ignored otherwise.
     * Main purpose of this method is to quickly open the generated file for manual review.
     * @param file file to be opened
     */
    private static void open(File file) {
        try {
            if (Desktop.desktopSupported && Desktop.desktop.isSupported(Desktop.Action.OPEN)) {
                Desktop.desktop.open(file)
                Thread.sleep(10000)
            }
        } catch (ignored) {
            // CI
        }
    }

}
