"use client";

import { useEffect, useRef, useState } from "react";

const WAKE_WORD_STORAGE_KEY = "jarvis-wake-word-enabled";
const ALWAYS_LISTENING_STORAGE_KEY = "jarvis-always-listening";
const SPEECH_OUTPUT_START_EVENT = "jarvis-speech-output-start";
const SPEECH_OUTPUT_END_EVENT = "jarvis-speech-output-end";
const COMMAND_SILENCE_MS = 1_500;

export type VoiceState =
  | "mic-off"
  | "push-to-talk"
  | "wake-word"
  | "always-listening"
  | "command";
type RecognitionMode = "off" | "push" | "wake" | "command";

interface SpeechRecognitionAlternativeLike {
  transcript: string;
}

interface SpeechRecognitionResultLike {
  readonly isFinal: boolean;
  readonly length: number;
  [index: number]: SpeechRecognitionAlternativeLike;
}

interface SpeechRecognitionResultListLike {
  readonly length: number;
  [index: number]: SpeechRecognitionResultLike;
}

interface SpeechRecognitionEventLike extends Event {
  readonly resultIndex: number;
  readonly results: SpeechRecognitionResultListLike;
}

interface SpeechRecognitionErrorEventLike extends Event {
  readonly error: string;
}

interface SpeechRecognitionLike {
  continuous: boolean;
  interimResults: boolean;
  lang: string;
  onresult: ((event: SpeechRecognitionEventLike) => void) | null;
  onerror: ((event: SpeechRecognitionErrorEventLike) => void) | null;
  onend: (() => void) | null;
  start(): void;
  stop(): void;
  abort(): void;
}

interface SpeechRecognitionConstructor {
  new (): SpeechRecognitionLike;
}

type SpeechRecognitionWindow = Window & {
  SpeechRecognition?: SpeechRecognitionConstructor;
  webkitSpeechRecognition?: SpeechRecognitionConstructor;
};

function joinTranscript(finalText: string, interimText: string) {
  return [finalText.trim(), interimText.trim()].filter(Boolean).join(" ");
}

export function useVoiceInput() {
  const [isSupported, setIsSupported] = useState(false);
  const [isListening, setIsListening] = useState(false);
  const [isWakeWordEnabled, setIsWakeWordEnabled] = useState(false);
  const [isAlwaysListening, setIsAlwaysListening] = useState(false);
  const [isMicOff, setIsMicOff] = useState(false);
  const [voiceState, setVoiceState] = useState<VoiceState>("push-to-talk");
  const [transcript, setTranscript] = useState("");
  const [error, setError] = useState<string | null>(null);
  const recognitionRef = useRef<SpeechRecognitionLike | null>(null);
  const modeRef = useRef<RecognitionMode>("off");
  const wakeEnabledRef = useRef(false);
  const alwaysListeningRef = useRef(false);
  const micOffRef = useRef(false);
  const fatalErrorRef = useRef(false);
  const outputSpeakingRef = useRef(false);
  const shouldRestartWakeRef = useRef(false);
  const restartTimerRef = useRef<ReturnType<typeof setTimeout> | null>(null);
  const commandTimerRef = useRef<ReturnType<typeof setTimeout> | null>(null);
  const finalTranscriptRef = useRef("");

  useEffect(() => {
    const speechWindow = window as SpeechRecognitionWindow;
    const Recognition =
      speechWindow.SpeechRecognition ?? speechWindow.webkitSpeechRecognition;

    if (!Recognition) return;

    const recognition = new Recognition();
    recognition.continuous = true;
    recognition.interimResults = true;
    recognition.lang = navigator.language || "en-US";

    const clearCommandTimer = () => {
      if (commandTimerRef.current) clearTimeout(commandTimerRef.current);
      commandTimerRef.current = null;
    };

    const beginRecognition = (mode: RecognitionMode) => {
      if (outputSpeakingRef.current || micOffRef.current || mode === "off") return;

      modeRef.current = mode;
      shouldRestartWakeRef.current =
        mode === "wake" && alwaysListeningRef.current;

      try {
        recognition.start();
        setVoiceState(
          mode === "wake"
            ? alwaysListeningRef.current
              ? "always-listening"
              : "wake-word"
            : "command",
        );
        setIsListening(mode === "push" || mode === "command");
      } catch {
        if (mode !== "wake") {
          setError("Voice recognition could not be started.");
          modeRef.current = "off";
          setVoiceState("push-to-talk");
          setIsListening(false);
        }
      }
    };

    const scheduleCommandCompletion = () => {
      clearCommandTimer();
      commandTimerRef.current = setTimeout(() => {
        if (modeRef.current !== "command") return;
        const shouldResumeWake =
          wakeEnabledRef.current && alwaysListeningRef.current;
        modeRef.current = shouldResumeWake ? "wake" : "off";
        shouldRestartWakeRef.current = shouldResumeWake;
        setIsListening(false);
        setVoiceState(shouldResumeWake ? "always-listening" : "push-to-talk");
        recognition.stop();
      }, COMMAND_SILENCE_MS);
    };

    recognition.onresult = (event) => {
      let interimText = "";

      for (let index = event.resultIndex; index < event.results.length; index++) {
        const result = event.results[index];
        const resultText = result[0]?.transcript.trim() ?? "";
        if (!resultText) continue;

        if (modeRef.current === "wake") {
          if (!result.isFinal) continue;

          const wakeMatch = resultText.match(/\bhey\s*,?\s*jarvis\b[\s,.:;!?-]*/i);
          if (!wakeMatch) continue;

          const commandText = resultText.slice(
            (wakeMatch.index ?? 0) + wakeMatch[0].length,
          ).trim();
          modeRef.current = "command";
          shouldRestartWakeRef.current = false;
          finalTranscriptRef.current = commandText;
          setTranscript(commandText);
          setVoiceState("command");
          setIsListening(true);
          setError(null);
          if (commandText) scheduleCommandCompletion();
          continue;
        }

        if (modeRef.current !== "push" && modeRef.current !== "command") {
          continue;
        }

        if (result.isFinal) {
          finalTranscriptRef.current = joinTranscript(
            finalTranscriptRef.current,
            resultText,
          );
        } else {
          interimText = joinTranscript(interimText, resultText);
        }
      }

      if (modeRef.current === "push" || modeRef.current === "command") {
        setTranscript(joinTranscript(finalTranscriptRef.current, interimText));
        if (modeRef.current === "command") scheduleCommandCompletion();
      }
    };

    recognition.onerror = (event) => {
      if (event.error === "aborted") {
        setIsListening(false);
        return;
      }

      const permissionDenied =
        event.error === "not-allowed" || event.error === "service-not-allowed";
      const unrecoverable =
        permissionDenied ||
        event.error === "audio-capture" ||
        event.error === "network";

      if (unrecoverable) {
        wakeEnabledRef.current = false;
        alwaysListeningRef.current = false;
        micOffRef.current = true;
        fatalErrorRef.current = true;
        shouldRestartWakeRef.current = false;
        modeRef.current = "off";
        setIsWakeWordEnabled(false);
        setIsAlwaysListening(false);
        setIsMicOff(true);
        setVoiceState("mic-off");
        try {
          localStorage.setItem(WAKE_WORD_STORAGE_KEY, "false");
          localStorage.setItem(ALWAYS_LISTENING_STORAGE_KEY, "false");
        } catch {
          // Browser storage may be unavailable.
        }
      }

      setError(
        permissionDenied
          ? "Microphone permission was denied."
          : event.error === "no-speech"
            ? null
            : "Voice recognition stopped unexpectedly.",
      );
      setIsListening(false);
    };

    recognition.onend = () => {
      setIsListening(false);

      if (
        wakeEnabledRef.current &&
        alwaysListeningRef.current &&
        !micOffRef.current &&
        !fatalErrorRef.current &&
        !outputSpeakingRef.current &&
        (shouldRestartWakeRef.current || modeRef.current === "wake")
      ) {
        modeRef.current = "wake";
        setVoiceState("always-listening");
        restartTimerRef.current = setTimeout(() => beginRecognition("wake"), 250);
      } else if (!wakeEnabledRef.current) {
        modeRef.current = "off";
        setVoiceState(micOffRef.current ? "mic-off" : "push-to-talk");
      } else {
        modeRef.current = "off";
        setVoiceState(micOffRef.current ? "mic-off" : "push-to-talk");
      }
    };

    const pauseForSpeechOutput = () => {
      outputSpeakingRef.current = true;
      if (modeRef.current === "wake" || modeRef.current === "command") {
        shouldRestartWakeRef.current = alwaysListeningRef.current;
        recognition.stop();
      }
    };

    const resumeAfterSpeechOutput = () => {
      outputSpeakingRef.current = false;
      if (
        wakeEnabledRef.current &&
        alwaysListeningRef.current &&
        !micOffRef.current &&
        !fatalErrorRef.current &&
        modeRef.current !== "push"
      ) {
        modeRef.current = "wake";
        beginRecognition("wake");
      }
    };

    window.addEventListener(SPEECH_OUTPUT_START_EVENT, pauseForSpeechOutput);
    window.addEventListener(SPEECH_OUTPUT_END_EVENT, resumeAfterSpeechOutput);
    recognitionRef.current = recognition;

    // Browser speech recognition is only known after client mount.
    // eslint-disable-next-line react-hooks/set-state-in-effect
    setIsSupported(true);

    try {
      const savedWakeWord =
        localStorage.getItem(WAKE_WORD_STORAGE_KEY) === "true";
      const savedAlwaysListening =
        savedWakeWord &&
        localStorage.getItem(ALWAYS_LISTENING_STORAGE_KEY) === "true";

      if (savedWakeWord) {
        wakeEnabledRef.current = true;
        setIsWakeWordEnabled(true);
      }

      if (savedAlwaysListening) {
        alwaysListeningRef.current = true;
        setIsAlwaysListening(true);
        beginRecognition("wake");
      }
    } catch {
      // Keep voice controls usable if browser storage is unavailable.
    }

    return () => {
      clearCommandTimer();
      if (restartTimerRef.current) clearTimeout(restartTimerRef.current);
      window.removeEventListener(SPEECH_OUTPUT_START_EVENT, pauseForSpeechOutput);
      window.removeEventListener(SPEECH_OUTPUT_END_EVENT, resumeAfterSpeechOutput);
      recognition.onresult = null;
      recognition.onerror = null;
      recognition.onend = null;
      recognition.abort();
      recognitionRef.current = null;
    };
  }, []);

  const startListening = () => {
    const recognition = recognitionRef.current;
    if (
      !recognition ||
      isListening ||
      wakeEnabledRef.current ||
      micOffRef.current
    ) {
      return;
    }

    finalTranscriptRef.current = "";
    modeRef.current = "push";
    setTranscript("");
    setError(null);

    try {
      recognition.start();
      setVoiceState("command");
      setIsListening(true);
    } catch {
      modeRef.current = "off";
      setError("Voice recognition could not be started.");
      setVoiceState("push-to-talk");
      setIsListening(false);
    }
  };

  const stopListening = () => {
    if (!recognitionRef.current || !isListening) return;
    const shouldResumeWake =
      wakeEnabledRef.current && alwaysListeningRef.current;
    modeRef.current = shouldResumeWake ? "wake" : "off";
    shouldRestartWakeRef.current = shouldResumeWake;
    recognitionRef.current.stop();
    setIsListening(false);
    setVoiceState(shouldResumeWake ? "always-listening" : "push-to-talk");
  };

  const toggleWakeWord = () => {
    const recognition = recognitionRef.current;
    if (!recognition || micOffRef.current) return;

    const nextEnabled = !wakeEnabledRef.current;
    wakeEnabledRef.current = nextEnabled;
    setIsWakeWordEnabled(nextEnabled);
    setError(null);

    try {
      localStorage.setItem(WAKE_WORD_STORAGE_KEY, String(nextEnabled));
    } catch {
      // Keep voice controls usable if browser storage is unavailable.
    }

    if (nextEnabled) {
      finalTranscriptRef.current = "";
      setTranscript("");
      modeRef.current = "wake";
      shouldRestartWakeRef.current = true;
      setVoiceState("wake-word");
      try {
        recognition.start();
      } catch {
        // An ending push-to-talk session will restart wake listening.
      }
    } else {
      alwaysListeningRef.current = false;
      setIsAlwaysListening(false);
      shouldRestartWakeRef.current = false;
      modeRef.current = "off";
      setVoiceState(micOffRef.current ? "mic-off" : "push-to-talk");
      setIsListening(false);
      if (restartTimerRef.current) clearTimeout(restartTimerRef.current);
      if (commandTimerRef.current) clearTimeout(commandTimerRef.current);
      try {
        localStorage.setItem(ALWAYS_LISTENING_STORAGE_KEY, "false");
      } catch {
        // Browser storage may be unavailable.
      }
      recognition.abort();
    }
  };

  const toggleAlwaysListening = () => {
    const recognition = recognitionRef.current;
    if (!recognition || micOffRef.current) return;

    const nextEnabled = !alwaysListeningRef.current;
    alwaysListeningRef.current = nextEnabled;
    setIsAlwaysListening(nextEnabled);
    setError(null);

    if (nextEnabled) {
      wakeEnabledRef.current = true;
      setIsWakeWordEnabled(true);
      modeRef.current = "wake";
      shouldRestartWakeRef.current = true;
      setVoiceState("always-listening");
      try {
        localStorage.setItem(WAKE_WORD_STORAGE_KEY, "true");
        localStorage.setItem(ALWAYS_LISTENING_STORAGE_KEY, "true");
        recognition.start();
      } catch {
        // An ending recognition session will restart when safe.
      }
    } else {
      shouldRestartWakeRef.current = false;
      modeRef.current = "off";
      setVoiceState("push-to-talk");
      setIsListening(false);
      if (restartTimerRef.current) clearTimeout(restartTimerRef.current);
      if (commandTimerRef.current) clearTimeout(commandTimerRef.current);
      try {
        localStorage.setItem(ALWAYS_LISTENING_STORAGE_KEY, "false");
      } catch {
        // Browser storage may be unavailable.
      }
      recognition.abort();
    }
  };

  const toggleMicOff = () => {
    const recognition = recognitionRef.current;
    if (!recognition) return;

    const nextMicOff = !micOffRef.current;
    micOffRef.current = nextMicOff;
    setIsMicOff(nextMicOff);
    setError(null);

    if (nextMicOff) {
      shouldRestartWakeRef.current = false;
      modeRef.current = "off";
      setIsListening(false);
      setVoiceState("mic-off");
      if (restartTimerRef.current) clearTimeout(restartTimerRef.current);
      if (commandTimerRef.current) clearTimeout(commandTimerRef.current);
      recognition.abort();
      return;
    }

    fatalErrorRef.current = false;
    if (wakeEnabledRef.current && alwaysListeningRef.current) {
      modeRef.current = "wake";
      shouldRestartWakeRef.current = true;
      setVoiceState("always-listening");
      try {
        recognition.start();
      } catch {
        // Recognition will remain off if the browser cannot restart it.
      }
    } else {
      setVoiceState("push-to-talk");
    }
  };

  const resetTranscript = () => {
    finalTranscriptRef.current = "";
    setTranscript("");
    setError(null);
  };

  return {
    isSupported,
    isListening,
    isWakeWordEnabled,
    isAlwaysListening,
    isMicOff,
    voiceState,
    transcript,
    error,
    startListening,
    stopListening,
    toggleWakeWord,
    toggleAlwaysListening,
    toggleMicOff,
    resetTranscript,
  };
}
