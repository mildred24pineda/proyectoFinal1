using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Net.Http;
using System.Text;
using System.Text.Json;
using System.Windows.Forms;
using A = DocumentFormat.OpenXml.Drawing;


namespace Proyecto_1_Programacion1
{ 
    public partial class Form1 : Form
    {
        private TextBox txtTema;
        private RichTextBox rtbResultado;
        private Button btnEnviar;
        private Button btnLimpiar;
        private Button btnExportarWord;
        private Button btnExportarPPT;

        public Form1()
        {
            InitializeComponent();
            this.Load += new EventHandler(Form1_Load);
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            this.Size = new Size(650, 400);
            this.Text = "Investigador Automático con AI";

            txtTema = new TextBox
            {
                Location = new Point(20, 20),
                Size = new Size(580, 30),
                BackColor = ColorTranslator.FromHtml("#F5F5F5")
            };
            this.Controls.Add(txtTema);

            rtbResultado = new RichTextBox
            {
                Location = new Point(20, 60),
                Size = new Size(580, 90),
                BackColor = ColorTranslator.FromHtml("#F5F5F5")
            };
            this.Controls.Add(rtbResultado);

            btnEnviar = new Button
            {
                Location = new Point(20, 165),
                Size = new Size(100, 40),
                Text = "Enviar",
                BackColor = System.Drawing.Color.LightGreen,
                ForeColor = System.Drawing.Color.Black
            };
            btnEnviar.Click += new EventHandler(btnEnviar_Click);
            this.Controls.Add(btnEnviar);

            btnLimpiar = new Button
            {
                Location = new Point(140, 165),
                Size = new Size(100, 40),
                Text = "Limpiar",
                BackColor = System.Drawing.Color.LightYellow,
                ForeColor = System.Drawing.Color.Black
            };
            btnLimpiar.Click += new EventHandler(btnLimpiar_Click);
            this.Controls.Add(btnLimpiar);

            btnExportarWord = new Button
            {
                Location = new Point(260, 165),
                Size = new Size(150, 40),
                Text = "Exportar Word",
                BackColor = System.Drawing.Color.LightBlue,
                ForeColor = System.Drawing.Color.Black
            };
            btnExportarWord.Click += new EventHandler(btnExportarWord_Click);
            this.Controls.Add(btnExportarWord);

            btnExportarPPT = new Button
            {
                Location = new Point(430, 165),
                Size = new Size(170, 40),
                Text = "Exportar PowerPoint",
                BackColor = System.Drawing.Color.Orange,
                ForeColor = System.Drawing.Color.Black
            };
            btnExportarPPT.Click += new EventHandler(btnExportarPPT_Click);
            this.Controls.Add(btnExportarPPT);
        }

        private async void btnEnviar_Click(object sender, EventArgs e)
        {
            string tema = txtTema.Text.Trim();
            if (string.IsNullOrEmpty(tema))
            {
                MessageBox.Show("Por favor, ingrese un tema de investigación.", "Atención", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try
            {
                string apiKey = "MI API DE GROQ"; // <-- Reemplaza con tu clave válida
                string apiUrl = "https://api.groq.com/openai/v1/chat/completions";

                using (HttpClient client = new HttpClient())
                {
                    client.DefaultRequestHeaders.Clear();
                    client.DefaultRequestHeaders.Add("Authorization", $"Bearer {apiKey}");

                    var body = new
                    {
                        model = "llama3-70b-8192",
                        messages = new[]
                        {
                            new { role = "system", content = "Eres un experto investigador académico. Responde de forma clara y precisa." },
                            new { role = "user", content = $"Haz una investigación clara y completa sobre el tema: {tema}" }
                        },
                        temperature = 0.7
                    };

                    var jsonBody = JsonSerializer.Serialize(body);
                    var content = new StringContent(jsonBody, Encoding.UTF8, "application/json");

                    var response = await client.PostAsync(apiUrl, content);
                    response.EnsureSuccessStatusCode();

                    string responseContent = await response.Content.ReadAsStringAsync();
                    var jsonDoc = JsonDocument.Parse(responseContent);

                    string resultado = jsonDoc.RootElement
                        .GetProperty("choices")[0]
                        .GetProperty("message")
                        .GetProperty("content")
                        .GetString();

                    rtbResultado.Text = resultado;
                    GuardarEnBaseDeDatos(tema, resultado);

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error al conectar con la API de Groq: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnLimpiar_Click(object sender, EventArgs e)
        {
            txtTema.Clear();
            rtbResultado.Clear();
        }

        /// Exportar a Word
        private void btnExportarWord_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(rtbResultado.Text))
            {
                MessageBox.Show("No hay contenido para exportar.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            SaveFileDialog saveFileDialog = new SaveFileDialog
            {
                Filter = "Documentos Word (*.docx)|*.docx",
                Title = "Guardar como documento Word",
                FileName = "Investigación.docx"
            };

            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    using (WordprocessingDocument wordDoc = WordprocessingDocument.Create(saveFileDialog.FileName, WordprocessingDocumentType.Document))
                    {
                        MainDocumentPart mainPart = wordDoc.AddMainDocumentPart();
                        mainPart.Document = new Document();
                        Body body = new Body();

                        // Crear bordes de página tipo "cuadro"
                        SectionProperties sectionProps = new SectionProperties();
                        PageBorders borders = new PageBorders()
                        {
                            Display = PageBorderDisplayValues.AllPages,
                            OffsetFrom = PageBorderOffsetValues.Page
                        };
                        borders.TopBorder = new TopBorder() { Val = BorderValues.Single, Size = 24, Color = "000000" };
                        borders.BottomBorder = new BottomBorder() { Val = BorderValues.Single, Size = 24, Color = "000000" };
                        borders.LeftBorder = new LeftBorder() { Val = BorderValues.Single, Size = 24, Color = "000000" };
                        borders.RightBorder = new RightBorder() { Val = BorderValues.Single, Size = 24, Color = "000000" };
                        sectionProps.Append(borders);

                        string[] lineas = rtbResultado.Text.Split(new[] { "\r\n", "\n" }, StringSplitOptions.None);

                        foreach (string linea in lineas)
                        {
                            string textoOriginal = linea.Trim();
                            string textoNormal = NormalizarTexto(textoOriginal);
                            bool esTitulo = textoNormal.StartsWith("Título:") || textoNormal.StartsWith("Subtítulo:");

                            // Crear propiedades comunes
                            RunProperties runProps = new RunProperties();
                            runProps.Append(new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" });
                            runProps.Append(new FontSize() { Val = "24" }); // 12 pt

                            // Si es título o subtítulo, aplica negrita
                            if (esTitulo)
                            {
                                runProps.Append(new Bold());
                            }

                            // Eliminar cualquier cursiva (por si acaso)
                            runProps.RemoveAllChildren<Italic>();
                            runProps.RemoveAllChildren<ItalicComplexScript>();
                            //runProps.RemoveAllChildren<ItalicEastAsian>();

                            // Crear texto plano sin cursiva ni estilos
                            Run run = new Run(runProps, new DocumentFormat.OpenXml.Math.Text(textoNormal) { Space = SpaceProcessingModeValues.Preserve });

                            Paragraph paragraph = new Paragraph(run);
                            body.Append(paragraph);
                        }

                        body.Append(sectionProps);
                        mainPart.Document.Append(body);
                        mainPart.Document.Save();
                    }

                    MessageBox.Show("Documento de Word guardado exitosamente.", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Ocurrió un error al exportar: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        // Función para normalizar caracteres matemáticos estilizados
        private string NormalizarTexto(string texto)
        {
            var normalizado = new StringBuilder();

            foreach (char c in texto)
            {
                int code = (int)c;

                // Letras minúsculas cursivas
                if (code >= 0x1D44E && code <= 0x1D467)
                    normalizado.Append((char)('a' + (code - 0x1D44E)));

                // Letras mayúsculas cursivas
                else if (code >= 0x1D434 && code <= 0x1D44D)
                    normalizado.Append((char)('A' + (code - 0x1D434)));

                // Números negrita
                else if (code >= 0x1D7CE && code <= 0x1D7D7)
                    normalizado.Append((char)('0' + (code - 0x1D7CE)));

                else
                    normalizado.Append(c);
            }

            return normalizado.ToString();
        }

        // Exportar a PowerPoint
        //private void btnExportarPPT_Click(object sender, EventArgs e)
        //{
        //    if (string.IsNullOrWhiteSpace(rtbResultado.Text))
        //    {
        //        MessageBox.Show("No hay contenido para exportar.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information);
        //        return;
        //    }

        //    SaveFileDialog saveFileDialog = new SaveFileDialog
        //    {
        //        Filter = "Archivos PowerPoint (*.pptx)|*.pptx",
        //        Title = "Guardar presentación de PowerPoint"
        //    };

        //    if (saveFileDialog.ShowDialog() == DialogResult.OK)
        //    {
        //        using (PresentationDocument presentationDoc = PresentationDocument.Create(saveFileDialog.FileName, PresentationDocumentType.Presentation))
        //        {
        //            PresentationPart presentationPart = presentationDoc.AddPresentationPart();
        //            presentationPart.Presentation = new Presentation();

        //            SlidePart slidePart = presentationPart.AddNewPart<SlidePart>();
        //            slidePart.Slide = new Slide(new CommonSlideData(new ShapeTree()));

        //            // Añadir elementos requeridos
        //            Slide slide = slidePart.Slide;
        //            ShapeTree shapeTree = slide.CommonSlideData.ShapeTree;

        //            // Requerido para el marcador de posición del slide
        //            shapeTree.Append(new NonVisualGroupShapeProperties(
        //                new NonVisualDrawingProperties() { Id = 1, Name = "" },
        //                new NonVisualGroupShapeDrawingProperties(),
        //                new ApplicationNonVisualDrawingProperties()));

        //            shapeTree.Append(new GroupShapeProperties(new A.TransformGroup()));

        //            // Crear cuadro de texto
        //            Shape shape = new Shape();

        //            shape.NonVisualShapeProperties = new NonVisualShapeProperties(
        //                new NonVisualDrawingProperties() { Id = 2, Name = "Contenido" },
        //                new NonVisualShapeDrawingProperties(new A.ShapeLocks() { NoGrouping = true }),
        //                new ApplicationNonVisualDrawingProperties(new PlaceholderShape()));

        //            shape.ShapeProperties = new ShapeProperties();

        //            shape.TextBody = new TextBody(
        //                new A.BodyProperties(),
        //                new A.ListStyle(),
        //                new A.Paragraph(new A.Run(new A.Text(rtbResultado.Text)))
        //            );

        //            shapeTree.Append(shape);

        //            // Asociar slide al slideIdList
        //            presentationPart.Presentation.Append(new SlideIdList(
        //                new SlideId() { Id = 256U, RelationshipId = presentationPart.GetIdOfPart(slidePart) }
        //            ));

        //            presentationPart.Presentation.Save();
        //        }

        //        MessageBox.Show("El contenido se exportó correctamente a PowerPoint.", "Exportar PowerPoint", MessageBoxButtons.OK, MessageBoxIcon.Information);
        //    }
        //}
        private void btnExportarPPT_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(rtbResultado.Text))
            {
                MessageBox.Show("No hay contenido para exportar.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            SaveFileDialog saveFileDialog = new SaveFileDialog
            {
                Filter = "Archivos PowerPoint (*.pptx)|*.pptx",
                Title = "Guardar presentación de PowerPoint"
            };

            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                using (PresentationDocument presentationDoc = PresentationDocument.Create(saveFileDialog.FileName, PresentationDocumentType.Presentation))
                {
                    PresentationPart presentationPart = presentationDoc.AddPresentationPart();
                    presentationPart.Presentation = new Presentation();
                    SlideIdList slideIdList = new SlideIdList();
                    uint slideId = 256;

                    string[] lineas = rtbResultado.Text.Split(new[] { "\r\n", "\n" }, StringSplitOptions.RemoveEmptyEntries);

                    foreach (string linea in lineas)
                    {
                        SlidePart slidePart = presentationPart.AddNewPart<SlidePart>();
                        slidePart.Slide = new Slide(new CommonSlideData(new ShapeTree()));
                        Slide slide = slidePart.Slide;
                        ShapeTree shapeTree = slide.CommonSlideData.ShapeTree;

                        // Obligatorio
                        shapeTree.Append(new NonVisualGroupShapeProperties(
                            new NonVisualDrawingProperties() { Id = 1, Name = "" },
                            new NonVisualGroupShapeDrawingProperties(),
                            new ApplicationNonVisualDrawingProperties()));

                        shapeTree.Append(new GroupShapeProperties(new A.TransformGroup()));

                        // Crear cuadro de texto
                        Shape shape = new Shape();

                        shape.NonVisualShapeProperties = new NonVisualShapeProperties(
                            new NonVisualDrawingProperties() { Id = 2, Name = "Contenido" },
                            new NonVisualShapeDrawingProperties(new A.ShapeLocks() { NoGrouping = true }),
                            new ApplicationNonVisualDrawingProperties());

                        // Centramos el cuadro de texto en la slide
                        shape.ShapeProperties = new ShapeProperties(
                            new A.Transform2D(
                                new A.Offset() { X = 2000000, Y = 1500000 }, // posición
                                new A.Extents() { Cx = 6000000, Cy = 3000000 } // tamaño
                            )
                        );

                        // Agregamos el texto con formato
                        A.Paragraph paragraph = new A.Paragraph(
                                 new A.Run(
                                     new A.RunProperties()
                                     {
                                         FontSize = 2400, // 24 pt
                                         Bold = true,
                                         Language = "es-ES"
                                     },
                                     new A.LatinFont() { Typeface = "Arial" }, // ✅ Agregamos tipo de letra
                                     new A.Text(linea.Trim())
                                 )
                             );

                        shape.TextBody = new TextBody(
                            new A.BodyProperties(),
                            new A.ListStyle(),
                            paragraph
                        );

                        shapeTree.Append(shape);

                        // Asignar ID a la diapositiva
                        SlideId sldId = new SlideId()
                        {
                            Id = slideId++,
                            RelationshipId = presentationPart.GetIdOfPart(slidePart)
                        };

                        slideIdList.Append(sldId);
                    }

                    // Guardar lista de slides
                    presentationPart.Presentation.Append(slideIdList);
                    presentationPart.Presentation.Save();
                }

                MessageBox.Show("El contenido se exportó correctamente a PowerPoint.", "Exportar PowerPoint", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void GuardarEnBaseDeDatos(string prompt, string resultado)
        {
            string connectionString = "Data Source=DESKTOP-IJ26VIG\\SQLEXPRESS01;Initial Catalog=ProyectoFinal1_DB;Integrated Security=True";

            string query = "INSERT INTO InformacionIA (Prompt, Resultado) VALUES (@Prompt, @Resultado)";

            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                using (SqlCommand command = new SqlCommand(query, connection))
                {
                    command.Parameters.AddWithValue("@Prompt", prompt ?? string.Empty); // Evitar nulls
                    command.Parameters.AddWithValue("@Resultado", resultado);

                    connection.Open();
                    command.ExecuteNonQuery();
                }
            }
        }

    }
}
