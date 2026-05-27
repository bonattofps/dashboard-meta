import com.sun.net.httpserver.HttpExchange;
import com.sun.net.httpserver.HttpServer;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.net.InetSocketAddress;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.Duration;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.concurrent.TimeUnit;

public class app {
    private static final int PORTA = 8080;
    private static final Path BASE = Path.of("").toAbsolutePath();
    private static final Path INDEX = BASE.resolve("index.html");
    private static final Path CSS = BASE.resolve("style.css");
    private static final Path PYTHON = BASE.resolve("calculos.py");
    private static final Path PLANILHA = BASE.resolve("INDICADORES DO SUPORTE DE 2026.xlsx");

    public static void main(String[] args) throws Exception {
        HttpServer servidor = HttpServer.create(new InetSocketAddress(PORTA), 0);
        servidor.createContext("/", app::pagina);
        servidor.createContext("/style.css", app::css);
        servidor.createContext("/api/dados", app::dados);
        servidor.createContext("/api/importar", app::importar);
        servidor.setExecutor(null);
        servidor.start();
        System.out.println("Dashboard em http://localhost:" + PORTA);
    }

    private static void pagina(HttpExchange exchange) throws IOException {
        if (!"GET".equalsIgnoreCase(exchange.getRequestMethod())) {
            enviarTexto(exchange, 405, "Metodo nao permitido");
            return;
        }

        enviarArquivo(exchange, INDEX, "text/html; charset=utf-8");
    }

    private static void css(HttpExchange exchange) throws IOException {
        if (!"GET".equalsIgnoreCase(exchange.getRequestMethod())) {
            enviarTexto(exchange, 405, "Metodo nao permitido");
            return;
        }

        enviarArquivo(exchange, CSS, "text/css; charset=utf-8");
    }

    private static void dados(HttpExchange exchange) throws IOException {
        if (!"GET".equalsIgnoreCase(exchange.getRequestMethod())) {
            enviarTexto(exchange, 405, "Metodo nao permitido");
            return;
        }

        try {
            String json = executarPython(List.of("--json"));
            enviar(exchange, 200, "application/json; charset=utf-8", json.getBytes(StandardCharsets.UTF_8));
        } catch (Exception erro) {
            enviarTexto(exchange, 500, erro.getMessage());
        }
    }

    private static void importar(HttpExchange exchange) throws IOException {
        if (!"POST".equalsIgnoreCase(exchange.getRequestMethod())) {
            enviarTexto(exchange, 405, "Metodo nao permitido");
            return;
        }

        try {
            byte[] corpo = lerBytes(exchange.getRequestBody());
            String contentType = exchange.getRequestHeaders().getFirst("Content-Type");
            byte[] planilha = extrairPlanilha(corpo, contentType);
            Files.write(PLANILHA, planilha);

            String json = executarPython(List.of("--json"));
            enviar(exchange, 200, "application/json; charset=utf-8", json.getBytes(StandardCharsets.UTF_8));
        } catch (Exception erro) {
            enviarTexto(exchange, 500, erro.getMessage());
        }
    }

    private static String executarPython(List<String> argumentos) throws IOException, InterruptedException {
        IOException ultimoErro = null;

        for (String comandoPython : List.of("python", "py")) {
            ArrayList<String> comando = new ArrayList<>();
            comando.add(comandoPython);
            comando.add(PYTHON.toString());
            comando.addAll(argumentos);

            ProcessBuilder builder = new ProcessBuilder(comando);
            builder.directory(BASE.toFile());
            builder.redirectErrorStream(true);

            try {
                Process processo = builder.start();
                String saida;
                try (InputStream input = processo.getInputStream()) {
                    saida = lerTudo(input);
                }

                boolean terminou = processo.waitFor(Duration.ofSeconds(90).toMillis(), TimeUnit.MILLISECONDS);
                if (!terminou) {
                    processo.destroyForcibly();
                    throw new IOException("Tempo excedido ao executar Python.");
                }

                if (processo.exitValue() != 0) {
                    throw new IOException(saida.trim());
                }

                return saida.trim();
            } catch (IOException erro) {
                ultimoErro = erro;
            }
        }

        throw ultimoErro == null ? new IOException("Python nao encontrado.") : ultimoErro;
    }

    private static String lerTudo(InputStream input) throws IOException {
        return new String(lerBytes(input), StandardCharsets.UTF_8);
    }

    private static byte[] lerBytes(InputStream input) throws IOException {
        ByteArrayOutputStream buffer = new ByteArrayOutputStream();
        input.transferTo(buffer);
        return buffer.toByteArray();
    }

    private static byte[] extrairPlanilha(byte[] corpo, String contentType) throws IOException {
        if (contentType == null || !contentType.contains("multipart/form-data")) {
            throw new IOException("Envie a planilha pelo campo planilha.");
        }

        String marcador = "boundary=";
        int indiceBoundary = contentType.indexOf(marcador);
        if (indiceBoundary < 0) {
            throw new IOException("Boundary multipart nao encontrado.");
        }

        String boundaryTexto = contentType.substring(indiceBoundary + marcador.length()).replace("\"", "").trim();
        byte[] boundary = ("--" + boundaryTexto).getBytes(StandardCharsets.ISO_8859_1);
        byte[] quebraDupla = "\r\n\r\n".getBytes(StandardCharsets.ISO_8859_1);

        int posicao = indiceDe(corpo, boundary, 0);
        while (posicao >= 0) {
            int inicioCabecalho = posicao + boundary.length;
            if (inicioCabecalho + 1 < corpo.length && corpo[inicioCabecalho] == '-' && corpo[inicioCabecalho + 1] == '-') {
                break;
            }
            if (inicioCabecalho + 1 < corpo.length && corpo[inicioCabecalho] == '\r' && corpo[inicioCabecalho + 1] == '\n') {
                inicioCabecalho += 2;
            }

            int fimCabecalho = indiceDe(corpo, quebraDupla, inicioCabecalho);
            if (fimCabecalho < 0) {
                break;
            }

            String cabecalho = new String(corpo, inicioCabecalho, fimCabecalho - inicioCabecalho, StandardCharsets.ISO_8859_1);
            int inicioDados = fimCabecalho + quebraDupla.length;
            int proximoBoundary = indiceDe(corpo, boundary, inicioDados);
            if (proximoBoundary < 0) {
                break;
            }

            int fimDados = proximoBoundary;
            if (fimDados >= 2 && corpo[fimDados - 2] == '\r' && corpo[fimDados - 1] == '\n') {
                fimDados -= 2;
            }

            if (cabecalho.contains("name=\"planilha\"") && cabecalho.contains("filename=")) {
                return Arrays.copyOfRange(corpo, inicioDados, fimDados);
            }

            posicao = proximoBoundary;
        }

        throw new IOException("Arquivo da planilha nao encontrado no upload.");
    }

    private static int indiceDe(byte[] origem, byte[] busca, int inicio) {
        for (int i = Math.max(inicio, 0); i <= origem.length - busca.length; i++) {
            boolean encontrou = true;
            for (int j = 0; j < busca.length; j++) {
                if (origem[i + j] != busca[j]) {
                    encontrou = false;
                    break;
                }
            }
            if (encontrou) {
                return i;
            }
        }
        return -1;
    }

    private static void enviarArquivo(HttpExchange exchange, Path caminho, String tipo) throws IOException {
        if (!Files.exists(caminho)) {
            enviarTexto(exchange, 404, "Arquivo nao encontrado");
            return;
        }

        enviar(exchange, 200, tipo, Files.readAllBytes(caminho));
    }

    private static void enviarTexto(HttpExchange exchange, int status, String texto) throws IOException {
        enviar(exchange, status, "text/plain; charset=utf-8", texto.getBytes(StandardCharsets.UTF_8));
    }

    private static void enviar(HttpExchange exchange, int status, String tipo, byte[] corpo) throws IOException {
        exchange.getResponseHeaders().add("Content-Type", tipo);
        exchange.sendResponseHeaders(status, corpo.length);
        try (OutputStream output = exchange.getResponseBody()) {
            output.write(corpo);
        }
    }
}
