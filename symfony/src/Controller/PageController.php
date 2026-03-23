<?php

namespace App\Controller;

use Symfony\Bundle\FrameworkBundle\Controller\AbstractController;
use Symfony\Component\HttpFoundation\Response;
use Symfony\Component\Routing\Attribute\Route;

class PageController extends AbstractController
{
    #[Route('/', name: 'dashboard', methods: ['GET'])]
    public function dashboard(): Response
    {
        return $this->render('pages/dashboard.html.twig');
    }

    #[Route('/generator-perioade', name: 'generator_perioade', methods: ['GET'])]
    public function generatorPerioade(): Response
    {
        return $this->render('pages/generator_perioade.html.twig');
    }

    #[Route('/word-to-excel', name: 'word_to_excel', methods: ['GET'])]
    public function wordToExcel(): Response
    {
        return $this->render('pages/word_converter.html.twig');
    }

    #[Route('/xml-formatter', name: 'xml_formatter', methods: ['GET'])]
    public function xmlFormatter(): Response
    {
        return $this->render('pages/xml_formatter.html.twig');
    }
}
