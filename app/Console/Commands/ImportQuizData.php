<?php

namespace App\Console\Commands;

use Illuminate\Console\Command;
use PhpOffice\PhpWord\IOFactory;
use App\Models\Quiz;
use Exception;

class ImportQuizData extends Command
{
    protected $signature = 'import:quiz-data {file}';
    protected $description = 'Import quiz data from a Word document';

    public function __construct()
    {
        parent::__construct();
    }

    public function handle()
    {
        $filePath = $this->argument('file');

        // Convert to absolute path
        $filePath = realpath($filePath);

        // Check if the file exists
        if (!$filePath || !file_exists($filePath)) {
            $this->error("File does not exist: {$filePath}");
            return;
        }

        try {
            // Load the Word document
            $phpWord = IOFactory::load($filePath);

            // Get the text from the document
            $text = '';
            foreach ($phpWord->getSections() as $section) {
                foreach ($section->getElements() as $element) {
                    if (method_exists($element, 'getText')) {
                        $text .= $element->getText() . "\n";
                    }
                }
            }

            // Split the text by questions
            $questions = explode('#END', $text);

            foreach ($questions as $questionText) {
                if (trim($questionText) == '') {
                    continue;
                }

                // Extract the question and options
                preg_match('/#Q\](.*)\n#a\](.*)\n#b\](.*)\n#c\](.*)\n#d\](.*)\n#R\]: (option\d)/', $questionText, $matches);

                if (count($matches) == 7) {
                    Quiz::create([
                        'question' => trim($matches[1]),
                        'option_a' => trim($matches[2]),
                        'option_b' => trim($matches[3]),
                        'option_c' => trim($matches[4]),
                        'option_d' => trim($matches[5]),
                        'correct_option' => trim($matches[6]),
                    ]);
                }
            }

            $this->info('Quiz data imported successfully.');

        } catch (Exception $e) {
            $this->error('Error loading the Word document: ' . $e->getMessage());
        }
    }
}
