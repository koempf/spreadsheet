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
package builders.dsl.spreadsheet.builder.api;

import java.util.function.Consumer;

public interface RowDefinition extends HasStyle {

    RowDefinition cell();
    RowDefinition cell(Object value);
    RowDefinition cell(Consumer<CellDefinition> cellDefinition);
    RowDefinition cell(int column, Consumer<CellDefinition> cellDefinition);
    RowDefinition cell(String column, Consumer<CellDefinition> cellDefinition);

    RowDefinition group(Consumer<RowDefinition> insideGroupDefinition);
    RowDefinition collapse(Consumer<RowDefinition> insideGroupDefinition);

    RowDefinition style(String name, Consumer<CellStyleDefinition> styleDefinition);
    RowDefinition styles(Iterable<String> names, Consumer<CellStyleDefinition> styleDefinition);
    RowDefinition style(Consumer<CellStyleDefinition> styleDefinition);
    RowDefinition style(String name);
    RowDefinition styles(String... names);
    RowDefinition styles(Iterable<String> names);

}
