package builders.dsl.spreadsheet.builder.data;

import builders.dsl.spreadsheet.builder.api.CellDefinition;
import builders.dsl.spreadsheet.builder.api.ImageCreator;

import java.io.*;
import java.net.URI;
import java.net.URISyntaxException;
import java.util.Base64;

public class ImageHelper implements ImageCreator {

    private final CellNode cellNode;
    private final String format;

    public ImageHelper(CellNode cellNode, String format) {
        this.cellNode = cellNode;
        this.format = format;
    }

    @Override
    public CellDefinition from(String fileOrUrl) {
        if (fileOrUrl.startsWith("file://")) {
            try (FileInputStream fis = new FileInputStream(new File(new URI(fileOrUrl)))) {
                return from(fis);
            } catch (URISyntaxException | IOException e) {
                throw new RuntimeException("Exception reading file: " + fileOrUrl, e);
            }
        }
        MapNode mapNode = new MapNode();
        mapNode.set("type", format);
        mapNode.set("url", fileOrUrl);
        cellNode.node.set("image", mapNode);
        return cellNode;
    }

    @Override
    public CellDefinition from(InputStream stream) {
        return from(streamToBytes(stream));
    }

    @Override
    public CellDefinition from(byte[] imageData) {
        return from("data:image/" + format + ";base64," + Base64.getEncoder().encodeToString(imageData));
    }

    private byte[] streamToBytes(InputStream is) {
        try {
            ByteArrayOutputStream buffer = new ByteArrayOutputStream();
            int nRead;
            byte[] data = new byte[1024];
            while ((nRead = is.read(data, 0, data.length)) != -1) {
                buffer.write(data, 0, nRead);
            }

            buffer.flush();
            return buffer.toByteArray();
        } catch (IOException e) {
            throw new RuntimeException("Exception reading image stream: " + is, e);
        }
    }
}
