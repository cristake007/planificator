<?php

namespace App\Controller;

use App\Service\PlanningService;
use Symfony\Bundle\FrameworkBundle\Controller\AbstractController;
use Symfony\Component\HttpFoundation\File\UploadedFile;
use Symfony\Component\HttpFoundation\JsonResponse;
use Symfony\Component\HttpFoundation\Request;
use Symfony\Component\HttpFoundation\Response;
use Symfony\Component\Routing\Attribute\Route;

class ApiController extends AbstractController
{
    public function __construct(private readonly PlanningService $planningService)
    {
    }

    #[Route('/generate_schedule', name: 'generate_schedule', methods: ['POST'])]
    public function generateSchedule(Request $request): JsonResponse
    {
        $file = $request->files->get('input_file');
        if (!$file instanceof UploadedFile) {
            return $this->json(['success' => false, 'error' => 'Input file is required.'], Response::HTTP_BAD_REQUEST);
        }

        try {
            $months = array_values(array_filter(array_map('trim', explode(',', (string) $request->request->get('months', '')))));
            $holidays = array_values(array_filter(array_map('trim', explode(',', (string) $request->request->get('holidays', '')))));
            $result = $this->planningService->generateSchedule(
                $file->getPathname(),
                $file->getClientOriginalName() ?: 'input.xlsx',
                (int) $request->request->get('year', date('Y')),
                (int) $request->request->get('randomness', 5),
                $months,
                $holidays,
            );

            $status = (int) ($result['status'] ?? 200);
            unset($result['status']);
            return $this->json($result, $status);
        } catch (\Throwable $e) {
            return $this->json(['success' => false, 'error' => $e->getMessage()], Response::HTTP_BAD_REQUEST);
        }
    }

    #[Route('/export_schedule', name: 'export_schedule', methods: ['POST'])]
    public function exportSchedule(Request $request): Response
    {
        try {
            $payload = json_decode($request->getContent(), true) ?? [];
            $result = $this->planningService->exportSchedule(
                $payload['schedule'] ?? [],
                (int) ($payload['year'] ?? date('Y')),
                $payload['holidays'] ?? [],
            );

            return new Response($result['bytes'], Response::HTTP_OK, [
                'Content-Type' => $result['content_type'],
                'Content-Disposition' => 'attachment; filename="' . $result['filename'] . '"',
            ]);
        } catch (\Throwable $e) {
            return $this->json(['success' => false, 'error' => $e->getMessage()], Response::HTTP_BAD_REQUEST);
        }
    }

    #[Route('/format-xml', name: 'format_xml', methods: ['POST'])]
    public function formatXml(Request $request): Response
    {
        $file = $request->files->get('input_file');
        if (!$file instanceof UploadedFile) {
            return $this->json(['success' => false, 'error' => 'No file provided'], Response::HTTP_BAD_REQUEST);
        }

        try {
            $result = $this->planningService->formatXml($file->getPathname(), $file->getClientOriginalName() ?: 'input.xlsx');
            return new Response($result['bytes'], Response::HTTP_OK, [
                'Content-Type' => $result['content_type'],
                'Content-Disposition' => 'attachment; filename="' . $result['filename'] . '"',
            ]);
        } catch (\Throwable $e) {
            return $this->json(['success' => false, 'error' => $e->getMessage()], Response::HTTP_BAD_REQUEST);
        }
    }

    #[Route('/convert_word', name: 'convert_word', methods: ['POST'])]
    public function convertWord(Request $request): Response
    {
        $word = $request->files->get('word_file');
        $permalinks = $request->files->get('permalinks_file');

        if (!$word instanceof UploadedFile || !$permalinks instanceof UploadedFile) {
            return $this->json(['success' => false, 'error' => 'Both files are required.'], Response::HTTP_BAD_REQUEST);
        }

        try {
            $result = $this->planningService->convertWord(
                $word->getPathname(),
                $permalinks->getPathname(),
                $permalinks->getClientOriginalName() ?: 'schedule.xlsx',
            );

            return new Response($result['bytes'], Response::HTTP_OK, [
                'Content-Type' => $result['content_type'],
                'Content-Disposition' => 'attachment; filename="' . $result['filename'] . '"',
            ]);
        } catch (\Throwable $e) {
            return $this->json(['success' => false, 'error' => $e->getMessage()], Response::HTTP_BAD_REQUEST);
        }
    }
}
