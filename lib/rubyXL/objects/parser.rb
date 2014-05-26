module RubyXL

  class ParsedThingNumeric
  end

  class ParsedThingInt
  end

  class ParsedThingString
  end

  class ParsedThingBool
  end

  class ParsedThingMissingArg
  end

  class ParsedThingError
  end

  class ParsedThingArray
  end

  class ParsedThingName
  end

  class ParsedThingNameX
  end

  class ParsedThingArea
  end

  class ParsedThingAreaErr
  end

  class ParsedThingArea3d
  end

  class ParsedThingAreaErr3d
  end

  class ParsedThingAreaN
  end

  class ParsedThingArea3dRel
  end

  class ParsedThingRef
  end

  class ParsedThingRefErr
  end

  class ParsedThingRef3d
  end

  class ParsedThingErr3d
  end

  class ParsedThingRefRel
  end

  class ParsedThingRef3dRel
  end

  class ParsedThingTable
  end

  class ParsedThingTableExt
  end

  class ParsedThingFunc < FunctionCall
  end

  class ParsedThingFuncVar < FunctionCall
  end

  class ParsedThingAdd < BinaryOperator
  end

  class ParsedThingSubtract < BinaryOperator
  end

  class ParsedThingMultiply < BinaryOperator
  end

  class ParsedThingDivide < BinaryOperator
  end

  class ParsedThingPower < BinaryOperator
  end

  class ParsedThingConcat < BinaryOperator
  end

  class ParsedThingLess < BinaryOperator
  end

  class ParsedThingLessEqual < BinaryOperator
  end

  class ParsedThingEqual < BinaryOperator
  end

  class ParsedThingGreaterEqual < BinaryOperator
  end

  class ParsedThingGreater < BinaryOperator
  end

  class ParsedThingNotEqual < BinaryOperator
  end

  class ParsedThingIntersect < BinaryOperator
  end

  class ParsedThingUnion < BinaryOperator
  end

  class ParsedThingRange < BinaryOperator
  end

  class ParsedThingUnaryPlus < UnaryOperator
  end

  class ParsedThingUnaryMinus < UnaryOperator
  end

  class ParsedThingPercent < UnaryOperator
  end

end