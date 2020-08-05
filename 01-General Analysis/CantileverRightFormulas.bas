Attribute VB_Name = "CantileverRightFormulas"

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' BSD 3-Clause License                                                            '''
'''                                                                                 '''
''' Copyright (c) 2020, open-struct-engineer                                        '''
''' All rights reserved.                                                            '''
'''                                                                                 '''
''' Redistribution and use in source and binary forms, with or without              '''
''' modification, are permitted provided that the following conditions are met:     '''
'''                                                                                 '''
''' 1. Redistributions of source code must retain the above copyright notice, this  '''
'''   list of conditions and the following disclaimer.                              '''
'''                                                                                 '''
''' 2. Redistributions in binary form must reproduce the above copyright notice,    '''
'''   this list of conditions and the following disclaimer in the documentation     '''
'''   and/or other materials provided with the distribution.                        '''
'''                                                                                 '''
''' 3. Neither the name of the copyright holder nor the names of its                '''
'''   contributors may be used to endorse or promote products derived from          '''
'''   this software without specific prior written permission.                      '''
'''                                                                                 '''
''' THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS"     '''
''' AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE       '''
''' IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE  '''
''' DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDER OR CONTRIBUTORS BE LIABLE    '''
''' FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL      '''
''' DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR      '''
''' SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER      '''
''' CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY,   '''
''' OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE   '''
''' OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.            '''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Function cant_right_initialSlope(slope As Double, L As Double, E As Double, I As Double, x As Double, result As Integer) As Variant

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''                                                                             '''
    '''     Function for a Starting slope on a right side cantilever                '''
    '''                                                                             '''
    '''     Important Note:                                                         '''
    '''     All inputs must have consistent units                                   '''
    '''                                                                             '''
    '''     Result key:                                                             '''
    '''     0 = Left Reaction                                                       '''
    '''     1 = Right Reaction                                                      '''
    '''     2 = Shear at x                                                          '''
    '''     3 = Moment at x                                                         '''
    '''     4 = Cross Section Rotation/Slope at x                                   '''
    '''     5 = Deflection at x                                                     '''
    '''     6 = Fixed End Moment at Left Support (clockwise positive)               '''
    '''     7 = Fixed End Moment at Right Support (clockwise positive)              '''
    '''                                                                             '''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''                                                                             '''
    '''     Sign Convention:                                                        '''
    '''     Loads applied in the (-)y direction are positive                        '''
    '''     Clockwise moments are positive                                          '''
    '''                                                                             '''
    '''     Reactions in the (+)y direction are positive                            '''
    '''                                                                             '''
    '''     Internal:                                                               '''
    '''     Shear is positive in the (+)y direction                                 '''
    '''     Moment is positive clockwise                                            '''
    '''     Cross Section Rotation/Slope is positive counter-clockwise              '''
    '''     Upward deflection is in the (+)y direction                              '''
    '''                                                                             '''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    
    ''''''''''''''''''''''''''''''
    ''  Variable Definitions    ''
    ''''''''''''''''''''''''''''''
    
    Dim rl As Double    'Left Reaction for a simple span beam
    Dim rr As Double    'Right Reaction foa smiple span beam
    Dim vx As Double    'Shear at x
    Dim mx As Double    'Moment at x
    Dim sx As Double    'Cross section Rotation/Slope at x
    Dim dx As Double    'Deflection at x
    Dim femL As Double  'Left End Fixed End Moment - Clockwise Positive
    Dim femR As Double  'Right End Fixed End Moment - Clockwise Positive
    
    ''''''''''''''''''''''''''''
    ''  Common Calculations   ''
    ''''''''''''''''''''''''''''
    
    'Support Reactions
    rl = 0
    rr = 0
    
    '''''''''''''''''''''''''''
    ''  Result Selection     ''
    '''''''''''''''''''''''''''
    
    If result = 0 Then
        
        'Left Support Reaction
        cant_right_initialSlope = rl
        
    ElseIf result = 1 Then
        
        'Right Support Reaction
        cant_right_initialSlope = rr
    
    ElseIf result = 2 Then
        
        'Shear at x
        vx = 0
  
        cant_right_initialSlope = vx
    
    ElseIf result = 3 Then
    
        'Moment at x
        mx = 0
        
        cant_right_initialSlope = mx
            
    ElseIf result = 4 Then
        
        'Cross Section Rotation/Slope at x
        sx = slope
        
        cant_right_initialSlope = sx
    
    ElseIf result = 5 Then
        
        'Deflection at x
        If 0 <= x And x <= L Then
            dx = (slope * x)
        Else
            dx = 0
        End If
        
        cant_right_initialSlope = dx
    
    ElseIf result = 6 Then
        'Fixed End Moment Left
        femL = 0
        
        cant_right_initialSlope = femL
    
    ElseIf result = 7 Then
    
        'Fixed End Moment Right
        femR = 0
        
        cant_right_initialSlope = femR
    
    Else
    
        cant_right_initialSlope = CVErr(xlErrNA)
            
    End If
    
End Function

Function cant_right_point_load(p As Double, a As Double, L As Double, E As Double, I As Double, x As Double, result As Integer) As Variant

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''                                                                             '''
    '''     Function for a Point Load anywhere on a right side cantilever           '''
    '''                                                                             '''
    '''     Important Note:                                                         '''
    '''     All inputs must have consistent units                                   '''
    '''                                                                             '''
    '''     Result key:                                                             '''
    '''     0 = Left Reaction                                                       '''
    '''     1 = Right Reaction                                                      '''
    '''     2 = Shear at x                                                          '''
    '''     3 = Moment at x                                                         '''
    '''     4 = Cross Section Rotation/Slope at x                                   '''
    '''     5 = Deflection at x                                                     '''
    '''     6 = Fixed End Moment at Left Support (clockwise positive)               '''
    '''     7 = Fixed End Moment at Right Support (clockwise positive)              '''
    '''                                                                             '''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''                                                                             '''
    '''     Sign Convention:                                                        '''
    '''     Loads applied in the (-)y direction are positive                        '''
    '''     Clockwise moments are positive                                          '''
    '''                                                                             '''
    '''     Reactions in the (+)y direction are positive                            '''
    '''                                                                             '''
    '''     Internal:                                                               '''
    '''     Shear is positive in the (+)y direction                                 '''
    '''     Moment is positive clockwise                                            '''
    '''     Cross Section Rotation/Slope is positive counter-clockwise              '''
    '''     Upward deflection is in the (+)y direction                              '''
    '''                                                                             '''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    
    ''''''''''''''''''''''''''''''
    ''  Variable Definitions    ''
    ''''''''''''''''''''''''''''''
    
    Dim b As Double     'Distance from load to free end
    Dim rl As Double    'Left Reaction for a simple span beam
    Dim rr As Double    'Right Reaction foa smiple span beam
    Dim c1 As Double    'Integration constant - C1
    Dim c2 As Double    'Integration constant - C2
    Dim c3 As Double    'Integration constant - C3
    Dim c4 As Double    'Integration constant - C4
    Dim vx As Double    'Shear at x
    Dim mx As Double    'Moment at x
    Dim sx As Double    'Cross section Rotation/Slope at x
    Dim dx As Double    'Deflection at x
    Dim femL As Double  'Left End Fixed End Moment - Clockwise Positive
    Dim femR As Double  'Right End Fixed End Moment - Clockwise Positive
    
    ''''''''''''''''''''''''''''
    ''  Input Error Handling  ''
    ''''''''''''''''''''''''''''
    If p = 0 Then
        'Capture case where load may be a place holder with no value
        'so return 0 instead of an error
        cant_right_point_load = 0
    Else
    
        If a > L Or a < 0 Then
            cant_right_point_load = " A , must be between 0 and L"
        
        Else
          ''''''''''''''''''''''''''''
          ''  Common Calculations   ''
          ''''''''''''''''''''''''''''
          
          b = L - a
          
          'Support Reactions
          rl = p
          rr = 0
          ml = -1 * p * a
          
          'Integration Constants
          c1 = 0
          c2 = 0
          c3 = (0.5 * rl * a * a) + (ml * a) + c1
          c4 = (-1 * c3 * a) + ((1 / 6) * rl * a * a * a) + (0.5 * ml * a * a) + (c1 * a) + c2
          
          '''''''''''''''''''''''''''
          ''  Result Selection     ''
          '''''''''''''''''''''''''''
          
          If result = 0 Then
              
              'Left Support Reaction
              cant_right_point_load = rl
              
          ElseIf result = 1 Then
              
              'Right Support Reaction
              cant_right_point_load = rr
          
          ElseIf result = 2 Then
              
              'Shear at x
              If 0 <= x And x <= a Then
                  If x = 0 And a = 0 Then
                      vx = 0
                  Else
                      vx = p
                  End If
              Else
                  vx = 0
              End If
        
              cant_right_point_load = vx
          
          ElseIf result = 3 Then
          
              'Moment at x
              If 0 <= x And x <= a Then
                  mx = (rl * x) + ml
              Else
                  mx = 0
              End If
              
              cant_right_point_load = mx
                  
          ElseIf result = 4 Then
              
              'Cross Section Rotation/Slope at x
              If 0 <= x And x <= a Then
                  sx = ((0.5 * rl * x * x) + (ml * x) + c1) / (E * I)
              ElseIf a < x And x <= L Then
                  sx = c3 / (E * I)
              Else
                  sx = 0
              End If
              
              cant_right_point_load = sx
          
          ElseIf result = 5 Then
              
              'Deflection at x
              If 0 <= x And x <= a Then
                  dx = (((1 / 6) * rl * x * x * x) + (0.5 * ml * x * x) + (c1 * x) + c2) / (E * I)
              ElseIf a < x And x <= L Then
                  dx = ((c3 * x) + c4) / (E * I)
              Else
                  dx = 0
              End If
              
              cant_right_point_load = dx
          
          ElseIf result = 6 Then
              'Fixed End Moment Left
              femL = ml
              
              cant_right_point_load = femL
          
          ElseIf result = 7 Then
          
              'Fixed End Moment Right
              femR = 0
              
              cant_right_point_load = femR
          
          Else
          
              cant_right_point_load = CVErr(xlErrNA)
                  
          End If
        End If
    End If
    
End Function

Function cant_right_point_moment(m As Double, a As Double, L As Double, E As Double, I As Double, x As Double, result As Integer) As Variant

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''                                                                             '''
    '''     Function for a Point Moment anywhere on a right side cantilever         '''
    '''                                                                             '''
    '''     Important Note:                                                         '''
    '''     All inputs must have consistent units                                   '''
    '''                                                                             '''
    '''     Result key:                                                             '''
    '''     0 = Left Reaction                                                       '''
    '''     1 = Right Reaction                                                      '''
    '''     2 = Shear at x                                                          '''
    '''     3 = Moment at x                                                         '''
    '''     4 = Cross Section Rotation/Slope at x                                   '''
    '''     5 = Deflection at x                                                     '''
    '''     6 = Fixed End Moment at Left Support (clockwise positive)               '''
    '''     7 = Fixed End Moment at Right Support (clockwise positive)              '''
    '''                                                                             '''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''                                                                             '''
    '''     Sign Convention:                                                        '''
    '''     Loads applied in the (-)y direction are positive                        '''
    '''     Clockwise moments are positive                                          '''
    '''                                                                             '''
    '''     Reactions in the (+)y direction are positive                            '''
    '''                                                                             '''
    '''     Internal:                                                               '''
    '''     Shear is positive in the (+)y direction                                 '''
    '''     Moment is positive clockwise                                            '''
    '''     Cross Section Rotation/Slope is positive counter-clockwise              '''
    '''     Upward deflection is in the (+)y direction                              '''
    '''                                                                             '''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    
    ''''''''''''''''''''''''''''''
    ''  Variable Definitions    ''
    ''''''''''''''''''''''''''''''
    
    Dim b As Double     'Distance from load to free end
    Dim rl As Double    'Left Reaction for a simple span beam
    Dim rr As Double    'Right Reaction foa smiple span beam
    Dim c1 As Double    'Integration constant - C1
    Dim c2 As Double    'Integration constant - C2
    Dim c3 As Double    'Integration constant - C3
    Dim c4 As Double    'Integration constant - C4
    Dim vx As Double    'Shear at x
    Dim mx As Double    'Moment at x
    Dim sx As Double    'Cross section Rotation/Slope at x
    Dim dx As Double    'Deflection at x
    Dim femL As Double  'Left End Fixed End Moment - Clockwise Positive
    Dim femR As Double  'Right End Fixed End Moment - Clockwise Positive
    
    ''''''''''''''''''''''''''''
    ''  Input Error Handling  ''
    ''''''''''''''''''''''''''''
    If m = 0 Then
        'Capture case where load may be a place holder with no value
        'so return 0 instead of an error
        cant_right_point_moment = 0
    
    Else
    
        If a > L Or a < 0 Then
            cant_right_point_moment = " A , must be between 0 and L"
        
        Else
            ''''''''''''''''''''''''''''
            ''  Common Calculations   ''
            ''''''''''''''''''''''''''''
            
            b = L - a
            
            'Support Reactions
            rl = 0
            rr = 0
            ml = -1 * m
            
            'Integration Constants
            c1 = 0
            c2 = 0
            c3 = (ml * a) + c1
            c4 = (0.5 * ml * a * a) + (c1 * a) + c2 - (c3 * a)
            
            '''''''''''''''''''''''''''
            ''  Result Selection     ''
            '''''''''''''''''''''''''''
            
            If result = 0 Then
                
                'Left Support Reaction
                cant_right_point_moment = rl
                
            ElseIf result = 1 Then
                
                'Right Support Reaction
                cant_right_point_moment = rr
            
            ElseIf result = 2 Then
                
                'Shear at x
                vx = 0
            
                cant_right_point_moment = vx
            
            ElseIf result = 3 Then
            
                'Moment at x
                If 0 <= x And x <= a Then
                    mx = ml
                Else
                    mx = 0
                End If
                
                cant_right_point_moment = mx
                    
            ElseIf result = 4 Then
                
                'Cross Section Rotation/Slope at x
                If 0 <= x And x <= a Then
                    sx = ((ml * x) + c1) / (E * I)
                ElseIf a < x And x <= L Then
                    sx = c3 / (E * I)
                Else
                    sx = 0
                End If
                
                cant_right_point_moment = sx
            
            ElseIf result = 5 Then
                
                'Deflection at x
                If 0 <= x And x <= a Then
                    dx = ((0.5 * ml * x * x) + (c1 * x) + c2) / (E * I)
                ElseIf a < x And x <= L Then
                    dx = ((c3 * x) + c4) / (E * I)
                Else
                    dx = 0
                End If
                
                cant_right_point_moment = dx
            
            ElseIf result = 6 Then
                'Fixed End Moment Left
                femL = ml
                
                cant_right_point_moment = femL
            
            ElseIf result = 7 Then
            
                'Fixed End Moment Right
                femR = 0
                
                cant_right_point_moment = femR
            
            Else
            
                cant_right_point_moment = CVErr(xlErrNA)
                    
            End If
        End If
    End If
    
End Function

Function cant_right_uniform_load(w As Double, a As Double, b As Double, L As Double, E As Double, I As Double, x As Double, result As Integer) As Variant

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''                                                                             '''
    '''     Function for a Uniform Distributed Load anywhere on                     '''
    '''     a right side cantilever                                                 '''
    '''                                                                             '''
    '''     Important Note:                                                         '''
    '''     All inputs must have consistent units                                   '''
    '''                                                                             '''
    '''     Result key:                                                             '''
    '''     0 = Left Reaction                                                       '''
    '''     1 = Right Reaction                                                      '''
    '''     2 = Shear at x                                                          '''
    '''     3 = Moment at x                                                         '''
    '''     4 = Cross Section Rotation/Slope at x                                   '''
    '''     5 = Deflection at x                                                     '''
    '''     6 = Fixed End Moment at Left Support (clockwise positive)               '''
    '''     7 = Fixed End Moment at Right Support (clockwise positive)              '''
    '''                                                                             '''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''                                                                             '''
    '''     Sign Convention:                                                        '''
    '''     Loads applied in the (-)y direction are positive                        '''
    '''     Clockwise moments are positive                                          '''
    '''                                                                             '''
    '''     Reactions in the (+)y direction are positive                            '''
    '''                                                                             '''
    '''     Internal:                                                               '''
    '''     Shear is positive in the (+)y direction                                 '''
    '''     Moment is positive clockwise                                            '''
    '''     Cross Section Rotation/Slope is positive counter-clockwise              '''
    '''     Upward deflection is in the (+)y direction                              '''
    '''                                                                             '''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    
    ''''''''''''''''''''''''''''''
    ''  Variable Definitions    ''
    ''''''''''''''''''''''''''''''
    
    Dim c As Double     'Width of load area
    Dim rl As Double    'Left Reaction for a simple span beam
    Dim rr As Double    'Right Reaction foa smiple span beam
    Dim c1 As Double    'Integration constant - C1
    Dim c2 As Double    'Integration constant - C2
    Dim c3 As Double    'Integration constant - C3
    Dim c4 As Double    'Integration constant - C4
    Dim c5 As Double    'Integration constant - C5
    Dim c6 As Double    'Integration constant - C6
    Dim w_tot As Double 'Area of load or equivalent point load
    Dim vx As Double    'Shear at x
    Dim mx As Double    'Moment at x
    Dim sx As Double    'Cross section Rotation/Slope at x
    Dim dx As Double    'Deflection at x
    Dim femL As Double  'Left End Fixed End Moment - Clockwise Positive
    Dim femR As Double  'Right End Fixed End Moment - Clockwise Positive
    
    ''''''''''''''''''''''''''''
    ''  Input Error Handling  ''
    ''''''''''''''''''''''''''''
    If w = 0 Then
        'Capture case where load may be a place holder with no value
        'so return 0 instead of an error
        cant_right_uniform_load = 0
    
    Else
    
        If a > L Or b > L Or a = b Or a < 0 Or b < 0 Then
            cant_right_uniform_load = "A,B cannot be equal and must be between 0 and L"
        ElseIf b < a Then
            cant_right_uniform_load = "B must be greater than A"
        Else
        
            ''''''''''''''''''''''''''''
            ''  Common Calculations   ''
            ''''''''''''''''''''''''''''
            
            c = b - a
            w_tot = w * c
            
            'Support Reactions
            rl = w_tot
            rr = 0
            ml = -1 * w_tot * (b - (c / 2))
            
            'Integration Constants
            c1 = 0
            c2 = 0
            c3 = 0
            c4 = (c1 * a) + c2 - (c3 * a)
            c5 = (0.5 * w_tot * b * b) + (ml * b) - ((1 / 6) * w * (b - a) * (b - a) * (b - a)) + c3
            c6 = ((1 / 6) * w_tot * b * b * b) + (0.5 * ml * b * b) - ((1 / 24) * w * (b - a) * (b - a) * (b - a) * (b - a)) + (c3 * b) + c4 - (c5 * b)
            
            '''''''''''''''''''''''''''
            ''  Result Selection     ''
            '''''''''''''''''''''''''''
            
            If result = 0 Then
                
                'Left Support Reaction
                cant_right_uniform_load = rl
                
            ElseIf result = 1 Then
                
                'Right Support Reaction
                cant_right_uniform_load = rr
            
            ElseIf result = 2 Then
                
                'Shear at x
                If 0 <= x And x <= a Then
                    vx = rl
                ElseIf a < x And x <= b Then
                    vx = rl - (w * (x - a))
                Else
                    vx = 0
                End If
            
                cant_right_uniform_load = vx
            
            ElseIf result = 3 Then
            
                'Moment at x
                If 0 <= x And x <= a Then
                    mx = (rl * x) + ml
                ElseIf a < x And x <= b Then
                    mx = (rl * x) + ml - (w * (x - a) * ((x - a) / 2))
                Else
                    mx = 0
                End If
                
                cant_right_uniform_load = mx
                    
            ElseIf result = 4 Then
                
                'Cross Section Rotation/Slope at x
                If 0 <= x And x <= a Then
                    sx = ((0.5 * rl * x * x) + (ml * x) + c1) / (E * I)
                ElseIf a < x And x <= b Then
                    sx = ((0.5 * rl * x * x) + (ml * x) - ((1 / 6) * w * (x - a) * (x - a) * (x - a)) + c3) / (E * I)
                ElseIf b < x And x <= L Then
                    sx = c5 / (E * I)
                Else
                    sx = 0
                End If
                
                cant_right_uniform_load = sx
            
            ElseIf result = 5 Then
                
                'Deflection at x
                If 0 <= x And x <= a Then
                    dx = (((1 / 6) * rl * x * x * x) + (0.5 * ml * x * x) + (c1 * x) + c2) / (E * I)
                ElseIf a < x And x <= b Then
                    dx = (((1 / 6) * rl * x * x * x) + (0.5 * ml * x * x) - ((1 / 24) * w * (x - a) * (x - a) * (x - a) * (x - a)) + (c3 * x) + c4) / (E * I)
                ElseIf b < x And x <= L Then
                    dx = ((c5 * x) + c6) / (E * I)
                Else
                    dx = 0
                End If
                
                cant_right_uniform_load = dx
            
            ElseIf result = 6 Then
                'Fixed End Moment Left
                femL = ml
                
                cant_right_uniform_load = femL
            
            ElseIf result = 7 Then
            
                'Fixed End Moment Right
                femR = 0
                
                cant_right_uniform_load = femR
            
            Else
            
                cant_right_uniform_load = CVErr(xlErrNA)
                    
            End If
        End If
    End If
    
End Function

Function cant_right_variable_load(w1 As Double, w2 As Double, a As Double, b As Double, L As Double, E As Double, I As Double, x As Double, result As Integer) As Variant

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''                                                                             '''
    '''     Function for a Variable Distributed Load anywhere on                    '''
    '''     a right side cantilever                                                 '''
    '''                                                                             '''
    '''     Important Note:                                                         '''
    '''     All inputs must have consistent units                                   '''
    '''                                                                             '''
    '''     Result key:                                                             '''
    '''     0 = Left Reaction                                                       '''
    '''     1 = Right Reaction                                                      '''
    '''     2 = Shear at x                                                          '''
    '''     3 = Moment at x                                                         '''
    '''     4 = Cross Section Rotation/Slope at x                                   '''
    '''     5 = Deflection at x                                                     '''
    '''     6 = Fixed End Moment at Left Support (clockwise positive)               '''
    '''     7 = Fixed End Moment at Right Support (clockwise positive)              '''
    '''                                                                             '''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''                                                                             '''
    '''     Sign Convention:                                                        '''
    '''     Loads applied in the (-)y direction are positive                        '''
    '''     Clockwise moments are positive                                          '''
    '''                                                                             '''
    '''     Reactions in the (+)y direction are positive                            '''
    '''                                                                             '''
    '''     Internal:                                                               '''
    '''     Shear is positive in the (+)y direction                                 '''
    '''     Moment is positive clockwise                                            '''
    '''     Cross Section Rotation/Slope is positive counter-clockwise              '''
    '''     Upward deflection is in the (+)y direction                              '''
    '''                                                                             '''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    
    ''''''''''''''''''''''''''''''
    ''  Variable Definitions    ''
    ''''''''''''''''''''''''''''''
    
    Dim c As Double     'Width of load area
    Dim rl As Double    'Left Reaction for a simple span beam
    Dim rr As Double    'Right Reaction foa smiple span beam
    Dim c1 As Double    'Integration constant - C1
    Dim c2 As Double    'Integration constant - C2
    Dim c3 As Double    'Integration constant - C3
    Dim c4 As Double    'Integration constant - C4
    Dim c5 As Double    'Integration constant - C5
    Dim c6 As Double    'Integration constant - C6
    Dim c7 As Double    'Integration constant - C7
    Dim w As Double     'Area of load or equivalent point load
    Dim s As Double     'Slope of load area
    Dim d As Double     'Center of load area
    Dim vx As Double    'Shear at x
    Dim mx As Double    'Moment at x
    Dim sx As Double    'Cross section Rotation/Slope at x
    Dim dx As Double    'Deflection at x
    Dim femL As Double  'Left End Fixed End Moment - Clockwise Positive
    Dim femR As Double  'Right End Fixed End Moment - Clockwise Positive
    
    ''''''''''''''''''''''''''''
    ''  Input Error Handling  ''
    ''''''''''''''''''''''''''''
    If w1 = 0 And w2 = 0 Then
        'Capture case where load may be a place holder with no value
        'so return 0 instead of an error
        cant_right_variable_load = 0
    
    Else
    
        If a > L Or b > L Or a = b Or a < 0 Or b < 0 Then
            cant_right_variable_load = "A,B cannot be equal and must be between 0 and L"
        ElseIf b < a Then
            cant_right_variable_load = "B must be greater than A"
        ElseIf Sgn(w1) <> Sgn(w2) And w1 <> 0 And w2 <> 0 Then
            cant_right_variable_load = "W1,W2 must have the same sign"
        Else
            
            ''''''''''''''''''''''''''''
            ''  Common Calculations   ''
            ''''''''''''''''''''''''''''
            
            c = b - a
            w = 0.5 * (w1 + w2) * c
            d = a + (((w1 + (2 * w2)) / (3 * (w2 + w1))) * c)
            s = (w1 - w2) / c
            
            'Support Reactions
            rl = w
            rr = 0
            ml = -1 * w * d
            
            'Integration Constants
            c1 = 0
            c2 = 0
            c3 = ml - ((1 / 6) * s * a * a * a) + (0.5 * (s * a + w1) * a * a) - (0.5 * (s * a + 2 * w1) * a * a)
            c4 = c1 - ((1 / 24) * s * a * a * a * a) + ((1# / 6#) * ((s * a) + w1) * a * a * a) - (0.25 * ((s * a) + (2 * w1)) * a * a * a) - (c3 * a) + (ml * a)
            c5 = (c1 * a) + c2 - (c4 * a) - ((1 / 120) * s * a * a * a * a * a) + ((1 / 24) * ((s * a) + w1) * a * a * a * a) - ((1 / 12) * ((s * a) + (2 * w1)) * a * a * a * a) + (0.5 * ml * a * a) - (0.5 * c3 * a * a)
            c6 = (0.5 * rl * b * b) + (c3 * b) + ((1 / 24) * s * b * b * b * b) - ((1 / 6) * ((s * a) + w1) * b * b * b) + (0.25 * ((s * a) + (2 * w1)) * a * b * b) + c4
            c7 = ((1 / 6) * rl * b * b * b) + (0.5 * c3 * b * b) + ((1 / 120) * s * b * b * b * b * b) - ((1 / 24) * ((s * a) + w1) * b * b * b * b) + ((1 / 12) * ((s * a) + (2 * w1)) * a * b * b * b) + (c4 * b) + c5 - (c6 * b)
            
            '''''''''''''''''''''''''''
            ''  Result Selection     ''
            '''''''''''''''''''''''''''
            
            If result = 0 Then
                
                'Left Support Reaction
                cant_right_variable_load = rl
                
            ElseIf result = 1 Then
                
                'Right Support Reaction
                cant_right_variable_load = rr
            
            ElseIf result = 2 Then
                
                'Shear at x
                If 0 <= x And x <= a Then
                    vx = rl
                ElseIf a < x And x <= b Then
                    vx = rl + (0.5 * s * x * x) - (x * ((s * a) + w1)) + (0.5 * a * ((s * a) + (2 * w1)))
                Else
                    vx = 0
                End If
            
                cant_right_variable_load = vx
            
            ElseIf result = 3 Then
            
                'Moment at x
                If 0 <= x And x <= a Then
                    mx = (rl * x) + ml
                ElseIf a < x And x <= b Then
                    mx = (rl * x) + c3 + ((1 / 6) * s * x * x * x) - (0.5 * ((s * a) + w1) * x * x) + (0.5 * ((s * a) + (2 * w1)) * a * x)
                Else
                    mx = 0
                End If
                
                cant_right_variable_load = mx
                    
            ElseIf result = 4 Then
                
                'Cross Section Rotation/Slope at x
                If 0 <= x And x <= a Then
                    sx = ((0.5 * rl * x * x) + (ml * x) + c1) / (E * I)
                ElseIf a < x And x <= b Then
                    sx = ((0.5 * rl * x * x) + (c3 * x) + ((1 / 24) * s * x * x * x * x) - ((1 / 6) * ((s * a) + w1) * x * x * x) + (0.25 * ((s * a) + (2 * w1)) * a * x * x) + c4) / (E * I)
                ElseIf b < x And x <= L Then
                    sx = c6 / (E * I)
                Else
                    sx = 0
                End If
                
                cant_right_variable_load = sx
            
            ElseIf result = 5 Then
                
                'Deflection at x
                If 0 <= x And x <= a Then
                    dx = (((1 / 6) * rl * x * x * x) + (0.5 * ml * x * x) + (c1 * x) + c2) / (E * I)
                ElseIf a < x And x <= b Then
                    dx = (((1 / 6) * rl * x * x * x) + (0.5 * c3 * x * x) + ((1 / 120) * s * x * x * x * x * x) - ((1 / 24) * ((s * a) + w1) * x * x * x * x) + ((1 / 12) * ((s * a) + (2 * w1)) * a * x * x * x) + (c4 * x) + c5) / (E * I)
                ElseIf b < x And x <= L Then
                    dx = ((c6 * x) + c7) / (E * I)
                Else
                    dx = 0
                End If
                
                cant_right_variable_load = dx
            
            ElseIf result = 6 Then
                'Fixed End Moment Left
                femL = ml
                
                cant_right_variable_load = femL
            
            ElseIf result = 7 Then
            
                'Fixed End Moment Right
                femR = 0
                
                cant_right_variable_load = femR
            
            Else
            
                cant_right_variable_load = CVErr(xlErrNA)
                    
            End If
        End If
    End If
End Function
