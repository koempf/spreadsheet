package builders.dsl.spreadsheet.onedrive;

import com.microsoft.graph.concurrency.ChunkedUploadProvider;
import com.microsoft.graph.concurrency.IProgressCallback;
import com.microsoft.graph.core.ClientException;
import com.microsoft.graph.models.extensions.DriveItem;
import com.microsoft.graph.models.extensions.DriveItemUploadableProperties;
import com.microsoft.graph.models.extensions.UploadSession;
import com.microsoft.graph.requests.extensions.ChunkedUploadRequest;
import com.microsoft.graph.requests.extensions.GraphServiceClient;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.Collections;
import java.util.concurrent.atomic.AtomicReference;
import java.util.function.Consumer;

public class OneDrive {

    private final GraphServiceClient client;

    public OneDrive(GraphServiceClient client) {
        this.client = client;
    }

    public DriveItem updateAndConvert(final String id, final String path, final String name, final Iterable<String> fields, final Consumer<OutputStream> withOutputStream) {
        try {
            DriveItemUploadableProperties item = new DriveItemUploadableProperties();
            item.name = name;
            UploadSession post = client.me().drive().root().itemWithPath(path).createUploadSession(item).buildRequest().post();

            ByteArrayOutputStream baos = new ByteArrayOutputStream();

            withOutputStream.accept(baos);

            byte[] buf = baos.toByteArray();
            ChunkedUploadProvider<DriveItem> provider = new ChunkedUploadProvider(post, client, new ByteArrayInputStream(buf), buf.length, DriveItem.class);

            AtomicReference<DriveItem> uploaded = new AtomicReference<>();

            provider.upload(new IProgressCallback<DriveItem>() {
                @Override
                public void progress(long current, long max) {
                    // ignored
                }

                @Override
                public void success(DriveItem driveItem) {
                    uploaded.set(driveItem);
                }

                @Override
                public void failure(ClientException ex) {
                    throw new IllegalArgumentException("Problems while uploading file " + name + " to path " + path, ex);
                }
            });
            return uploaded.get();
        } catch (IOException e) {
            throw new IllegalArgumentException("Problems while uploading file " + name + " to path " + path, e);
        }
    }

}
