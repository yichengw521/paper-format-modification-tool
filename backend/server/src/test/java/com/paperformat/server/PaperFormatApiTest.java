package com.paperformat.server;

import org.junit.jupiter.api.Test;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.test.autoconfigure.web.servlet.AutoConfigureMockMvc;
import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.mock.web.MockMultipartFile;
import org.springframework.test.web.servlet.MockMvc;

import static org.springframework.test.web.servlet.request.MockMvcRequestBuilders.get;
import static org.springframework.test.web.servlet.request.MockMvcRequestBuilders.multipart;
import static org.springframework.test.web.servlet.result.MockMvcResultMatchers.jsonPath;
import static org.springframework.test.web.servlet.result.MockMvcResultMatchers.status;

@SpringBootTest(properties = "paper-format.storage-root=${java.io.tmpdir}/paper-format-server-test")
@AutoConfigureMockMvc
class PaperFormatApiTest {
    @Autowired
    private MockMvc mockMvc;

    @Test
    void healthEndpointIsAvailable() throws Exception {
        mockMvc.perform(get("/api/v1/health"))
                .andExpect(status().isOk())
                .andExpect(jsonPath("$.status").value("UP"));
    }

    @Test
    void missingTaskReturnsNotFound() throws Exception {
        mockMvc.perform(get("/api/v1/tasks/00000000-0000-0000-0000-000000000000"))
                .andExpect(status().isNotFound());
    }

    @Test
    void rejectsNonDocxUploads() throws Exception {
        MockMultipartFile template = new MockMultipartFile(
                "template", "template.txt", "text/plain", "not a docx".getBytes());
        MockMultipartFile document = new MockMultipartFile(
                "document", "document.docx", "application/octet-stream", "not a docx".getBytes());

        mockMvc.perform(multipart("/api/v1/tasks").file(template).file(document))
                .andExpect(status().isBadRequest())
                .andExpect(jsonPath("$.detail").value("template must be a .docx file."));
    }
}
